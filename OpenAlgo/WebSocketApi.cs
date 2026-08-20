using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.IO;

namespace OpenAlgo
{
    /// <summary>
    /// Manages WebSocket connections for real-time market data streaming from OpenAlgo.
    ///
    /// Data reaches Excel by a push model: every streaming cell registers a listener for
    /// its own topic ("symbol|exchange|mode") and is woken only when a tick for that topic
    /// arrives. Nothing in this class ever asks Excel to recalculate, which is what kept
    /// the application permanently busy and broke copy and paste in earlier versions.
    /// </summary>
    public class WebSocketManager
    {
        private static WebSocketManager? _instance;
        private static readonly object _lock = new object();

        private ClientWebSocket? _webSocket;
        private CancellationTokenSource? _cancellationTokenSource;
        private Task? _receiveTask;

        // Store real-time data for each subscription
        private readonly ConcurrentDictionary<string, JObject> _marketData = new();

        // Store subscription status
        private readonly ConcurrentDictionary<string, SubscriptionInfo> _subscriptions = new();

        // Track manually unsubscribed symbols to prevent auto-resubscribe
        private readonly ConcurrentDictionary<string, DateTime> _manuallyUnsubscribed = new();

        // Track in-progress auto-subscribe attempts to prevent duplicate background tasks
        private readonly ConcurrentDictionary<string, bool> _autoSubscribeInProgress = new();

        // Depth level last requested for a topic, so a reconnect can restore it
        private readonly ConcurrentDictionary<string, int> _requestedDepth = new();

        // Maps our topic key to the topic string the server used, when they differ.
        // Ticks are cached under both so either spelling can be read back, and the
        // alias has to be recorded to be evictable on unsubscribe.
        private readonly ConcurrentDictionary<string, string> _topicAliases = new();

        // Live cell listeners per topic. The inner dictionary is keyed by listener id.
        private readonly ConcurrentDictionary<string, ConcurrentDictionary<long, Action<JObject>>> _listeners = new();
        private long _listenerSequence;

        // Per topic throttle state: last push time, retained tick and trailing flush timer
        private readonly ConcurrentDictionary<string, TopicThrottle> _throttles = new();

        // Topics already reported as receiving data without an active subscription,
        // so the log does not grow once per tick in that state
        private readonly ConcurrentDictionary<string, bool> _warnedUnsubscribed = new();

        // Order updates, newest first
        private readonly List<JObject> _orderUpdates = new();
        private readonly object _orderLock = new object();
        private const int MaxOrderUpdates = 200;
        private bool _ordersSubscribed;
        private bool _ordersManuallyUnsubscribed;
        private bool _orderSubscribeInProgress;

        // Pending request/response exchanges keyed by the expected response type
        private readonly ConcurrentDictionary<string, TaskCompletionSource<JObject>> _pendingRequests = new();

        private bool _isAuthenticated = false;
        // Guarded with Interlocked rather than a plain bool: ConnectAsync is reached
        // concurrently from every auto-subscribing cell on a sheet, and a check-then-set
        // race let two callers each build a socket, each start a receive loop, and leak
        // whichever one lost the assignment.
        private int _connectingFlag;

        private bool _isConnecting => Volatile.Read(ref _connectingFlag) != 0;
        private DateTime _lastPingTime = DateTime.MinValue;

        // WebSocket configuration
        private string _wsUrl = "ws://127.0.0.1:8765";

        // Authentication confirmation tracking
        private TaskCompletionSource<bool>? _authTcs;

        // Subscription confirmation tracking
        private readonly ConcurrentDictionary<string, TaskCompletionSource<bool>> _pendingSubscriptions = new();

        /// <summary>
        /// Pseudo topic used by the account wide order update stream.
        /// </summary>
        public const string OrdersTopic = "__orders__";

        /// <summary>
        /// Message type used for status notifications pushed to cells, for example
        /// "Unsubscribed" or "Disconnected". Real ticks carry "market_data".
        /// </summary>
        public const string StatusMessageType = "openalgo_status";

        // Logging
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OpenAlgo",
            "websocket.log"
        );

        // Backward compatibility: Set to true if backend doesn't send confirmation messages
        private bool _legacyMode = true;  // Default to true for backward compatibility

        public static WebSocketManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= new WebSocketManager();
                    }
                }
                return _instance;
            }
        }

        private WebSocketManager()
        {
            _wsUrl = string.IsNullOrWhiteSpace(OpenAlgoConfig.WebSocketUrl)
                ? "ws://127.0.0.1:8765"
                : OpenAlgoConfig.WebSocketUrl;
        }

        /// <summary>
        /// Logs messages to a file for debugging
        /// </summary>
        private void Log(string level, string message)
        {
            try
            {
                var logDir = Path.GetDirectoryName(LogFilePath);
                if (logDir != null && !Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] {message}\n";
                File.AppendAllText(LogFilePath, logEntry);
            }
            catch
            {
                // Silently fail if logging fails
            }
        }

        /// <summary>
        /// Sets the WebSocket URL for connection
        /// </summary>
        public void SetWebSocketUrl(string url)
        {
            if (!string.IsNullOrWhiteSpace(url))
                _wsUrl = url;
        }

        /// <summary>
        /// Gets the URL the next connection attempt will use.
        /// </summary>
        public string GetWebSocketUrl() => _wsUrl;

        /// <summary>
        /// Gets the current connection state
        /// </summary>
        public string GetConnectionState()
        {
            if (_webSocket == null) return "Disconnected";
            return _webSocket.State.ToString();
        }

        /// <summary>
        /// True once the server has accepted the API key.
        /// </summary>
        public bool IsAuthenticated => _isAuthenticated;

        // ------------------------------------------------------------------
        // Push plumbing
        // ------------------------------------------------------------------

        /// <summary>
        /// Builds the topic name a streaming cell listens on.
        /// </summary>
        public static string TopicFor(string symbol, string exchange, int mode) =>
            $"{symbol}|{exchange}|{mode}";

        /// <summary>
        /// Registers a listener for a topic. The listener runs on the WebSocket receive
        /// thread and must not throw. Dispose the returned handle to detach it.
        /// </summary>
        public IDisposable AddListener(string topic, Action<JObject> onUpdate)
        {
            if (string.IsNullOrEmpty(topic) || onUpdate == null)
                return new ListenerHandle(null, "", 0);

            long id = Interlocked.Increment(ref _listenerSequence);
            var bag = _listeners.GetOrAdd(topic, _ => new ConcurrentDictionary<long, Action<JObject>>());
            bag[id] = onUpdate;
            return new ListenerHandle(this, topic, id);
        }

        private void RemoveListener(string topic, long id)
        {
            if (_listeners.TryGetValue(topic, out var bag))
            {
                bag.TryRemove(id, out _);

                // Nothing is watching this topic any more, so a trailing flush would
                // fire into a dead cell. Drop it.
                if (bag.IsEmpty)
                    CancelThrottle(topic);
            }
        }

        /// <summary>
        /// Handle returned by AddListener. Disposing detaches the callback.
        /// </summary>
        private sealed class ListenerHandle : IDisposable
        {
            private WebSocketManager? _owner;
            private readonly string _topic;
            private readonly long _id;

            public ListenerHandle(WebSocketManager? owner, string topic, long id)
            {
                _owner = owner;
                _topic = topic;
                _id = id;
            }

            public void Dispose()
            {
                var owner = Interlocked.Exchange(ref _owner, null);
                owner?.RemoveListener(_topic, _id);
            }
        }

        /// <summary>
        /// Throttle state for one topic: when it last pushed, the newest tick that the
        /// throttle held back, and the one-shot timer that will release it.
        /// </summary>
        private sealed class TopicThrottle : IDisposable
        {
            /// <summary>Guards every field below.</summary>
            public readonly object Gate = new object();

            /// <summary>Environment.TickCount64 at the last push, valid only when HasPushed.</summary>
            public long LastPushedTicks;

            /// <summary>False until the first push, so the first tick is never delayed.</summary>
            public bool HasPushed;

            /// <summary>Newest tick held back by the throttle, or null when nothing is pending.</summary>
            public JObject? Pending;

            /// <summary>At most one outstanding trailing flush per topic.</summary>
            public System.Threading.Timer? FlushTimer;

            /// <summary>
            /// Set once the topic is torn down. A flush that is already running checks
            /// this before pushing, so a retained tick cannot land on a cell that has
            /// since been unsubscribed or disconnected.
            /// </summary>
            public volatile bool Cancelled;

            public void Dispose()
            {
                lock (Gate)
                {
                    Cancelled = true;
                    Pending = null;
                    var timer = FlushTimer;
                    FlushTimer = null;
                    timer?.Dispose();
                }
            }
        }

        /// <summary>
        /// Pushes a payload to every listener on a topic, subject to the per topic
        /// throttle in OpenAlgoConfig.StreamThrottleMs. A throttle of 0 pushes every tick.
        ///
        /// The throttle is leading plus trailing, and the invariant it guarantees is:
        /// the last value the server sent for a topic always reaches the cell, at most
        /// StreamThrottleMs late. A tick that arrives inside the window is retained as
        /// the pending value and released by a one-shot trailing flush when the window
        /// expires; a newer tick replaces the pending value rather than queueing a
        /// second flush. Do not reduce this to a leading edge only throttle: dropping
        /// the held tick outright leaves the final print of an illiquid strike, or the
        /// closing price of any instrument, permanently stale in the sheet.
        ///
        /// The flush calls OnNext on the topics that have a newer value and nothing
        /// else. It never asks Excel to recalculate, which is what made issue #4 lock
        /// up the application.
        /// </summary>
        private void NotifyTopic(string topic, JObject payload, bool bypassThrottle = false)
        {
            if (!_listeners.TryGetValue(topic, out var bag) || bag.IsEmpty)
                return;

            int throttleMs = OpenAlgoConfig.StreamThrottleMs;

            if (bypassThrottle || throttleMs <= 0)
            {
                // Pushing now supersedes anything the throttle was holding.
                if (_throttles.TryGetValue(topic, out var current))
                {
                    lock (current.Gate)
                    {
                        current.Pending = null;
                        current.HasPushed = true;
                        current.LastPushedTicks = Environment.TickCount64;
                        var stale = current.FlushTimer;
                        current.FlushTimer = null;
                        stale?.Dispose();
                    }
                }

                Dispatch(topic, payload);
                return;
            }

            var state = _throttles.GetOrAdd(topic, _ => new TopicThrottle());
            long now = Environment.TickCount64;

            lock (state.Gate)
            {
                if (state.HasPushed && now - state.LastPushedTicks < throttleMs)
                {
                    // Inside the window: retain the newest tick and make sure exactly
                    // one trailing flush is outstanding for this topic.
                    state.Pending = payload;

                    if (state.FlushTimer == null)
                    {
                        long dueIn = throttleMs - (now - state.LastPushedTicks);
                        if (dueIn < 1)
                            dueIn = 1;

                        state.FlushTimer = new System.Threading.Timer(
                            _ => FlushTopic(topic), null, dueIn, System.Threading.Timeout.Infinite);
                    }
                    return;
                }

                // Leading edge: push straight away.
                state.HasPushed = true;
                state.LastPushedTicks = now;
                state.Pending = null;
            }

            Dispatch(topic, payload);
        }

        /// <summary>
        /// Releases the tick the throttle held back, unless a newer one has already
        /// been pushed in the meantime.
        /// </summary>
        private void FlushTopic(string topic)
        {
            try
            {
                if (!_throttles.TryGetValue(topic, out var state))
                    return;

                JObject? payload;
                lock (state.Gate)
                {
                    var timer = state.FlushTimer;
                    state.FlushTimer = null;
                    timer?.Dispose();

                    if (state.Cancelled)
                        return;

                    payload = state.Pending;
                    state.Pending = null;

                    if (payload != null)
                    {
                        state.HasPushed = true;
                        state.LastPushedTicks = Environment.TickCount64;
                    }
                }

                // Re-checked outside the lock: the topic may have been torn down while
                // this flush was waiting, and a stale price must not overwrite the
                // "Unsubscribed" or "Disconnected" status that replaced it.
                if (payload != null && !state.Cancelled)
                    Dispatch(topic, payload);
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Trailing flush for {topic} failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Hands a payload to every listener on a topic. Never throws.
        /// </summary>
        private void Dispatch(string topic, JObject payload)
        {
            if (!_listeners.TryGetValue(topic, out var bag) || bag.IsEmpty)
                return;

            foreach (var listener in bag.Values)
            {
                try
                {
                    listener(payload);
                }
                catch (Exception ex)
                {
                    // A misbehaving cell must never take down the receive loop.
                    Log("ERROR", $"Listener for {topic} threw: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Drops the throttle state for a topic, including any outstanding trailing
        /// flush. Used on unsubscribe, on disconnect and when the last cell detaches.
        /// </summary>
        private void CancelThrottle(string topic)
        {
            if (_throttles.TryRemove(topic, out var state))
                state.Dispose();
        }

        /// <summary>
        /// Drops every throttle state and any outstanding trailing flush.
        /// </summary>
        private void CancelAllThrottles()
        {
            foreach (var topic in _throttles.Keys.ToArray())
                CancelThrottle(topic);
        }

        /// <summary>
        /// Pushes a short status string, such as "Unsubscribed", to one topic.
        /// Any tick the throttle was holding for that topic is discarded: the status
        /// is the newer truth, and the pending value must not overwrite it later.
        /// </summary>
        private void NotifyStatus(string topic, string message)
        {
            var payload = new JObject
            {
                ["type"] = StatusMessageType,
                ["message"] = message
            };
            CancelThrottle(topic);
            NotifyTopic(topic, payload, bypassThrottle: true);
        }

        /// <summary>
        /// Pushes a status string to every topic that currently has listeners.
        /// </summary>
        private void NotifyAllStatus(string message)
        {
            foreach (var topic in _listeners.Keys.ToArray())
                NotifyStatus(topic, message);
        }

        /// <summary>
        /// Re-subscribes every topic that still has a live cell listening on it.
        /// Called after a successful connect so cells recover from a reconnect.
        /// </summary>
        private void ResubscribeActiveTopics()
        {
            foreach (var entry in _listeners.ToArray())
            {
                if (entry.Value.IsEmpty)
                    continue;

                string topic = entry.Key;

                if (topic == OrdersTopic)
                {
                    if (!_ordersManuallyUnsubscribed)
                        StartAutoSubscribeOrders();
                    continue;
                }

                if (_manuallyUnsubscribed.ContainsKey(topic) || _subscriptions.ContainsKey(topic))
                    continue;

                var parts = topic.Split('|');
                if (parts.Length != 3 || !int.TryParse(parts[2], out int mode))
                    continue;

                int? depth = _requestedDepth.TryGetValue(topic, out int level) ? level : (int?)null;
                StartAutoSubscribe(parts[0], parts[1], mode, depth);
            }
        }

        // ------------------------------------------------------------------
        // Connection
        // ------------------------------------------------------------------

        /// <summary>
        /// Connects to the WebSocket server and authenticates
        /// </summary>
        public async Task<string> ConnectAsync()
        {
            try
            {
                if (_webSocket?.State == WebSocketState.Open)
                {
                    Log("INFO", "Already connected");
                    return "Already connected";
                }

                // Claim the connect slot atomically. Only the caller that flips 0 to 1
                // builds a socket; everyone else is told a connect is already running.
                if (Interlocked.CompareExchange(ref _connectingFlag, 1, 0) != 0)
                {
                    Log("WARN", "Connection already in progress");
                    return "Connection already in progress";
                }

                Log("INFO", $"Connecting to {_wsUrl}");

                // Release whatever a previous failed or closed attempt left behind before
                // overwriting the fields, otherwise the old socket, its token source and
                // its receive loop stay alive and keep mutating the cache.
                DisposeSocket();

                _webSocket = new ClientWebSocket();
                _cancellationTokenSource = new CancellationTokenSource();

                await _webSocket.ConnectAsync(new Uri(_wsUrl), _cancellationTokenSource.Token);
                Log("INFO", "WebSocket connected successfully");

                // Start receiving messages
                _receiveTask = Task.Run(() => ReceiveMessagesAsync(_cancellationTokenSource.Token));

                // Wait briefly for API key if not set (handles recalculation order race
                // where oa_ws_connect runs before oa_api sets the key)
                if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                {
                    Log("INFO", "Waiting for API key to be set...");
                    for (int i = 0; i < 20 && string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey); i++)
                    {
                        await Task.Delay(100);
                    }
                }

                // Authenticate
                var authResult = await AuthenticateAsync();

                if (!_isAuthenticated)
                {
                    // Clean up to prevent half-connected dead state
                    Log("WARN", "Authentication failed, cleaning up WebSocket connection");
                    try
                    {
                        _cancellationTokenSource?.Cancel();
                        if (_webSocket?.State == WebSocketState.Open)
                            await _webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "Auth failed", CancellationToken.None);
                    }
                    catch { }
                    _webSocket = null;
                    Volatile.Write(ref _connectingFlag, 0);
                    return authResult;
                }

                Volatile.Write(ref _connectingFlag, 0);

                // Restore anything a live cell is still watching
                ResubscribeActiveTopics();

                return authResult;
            }
            catch (Exception ex)
            {
                Volatile.Write(ref _connectingFlag, 0);
                Log("ERROR", $"Connection failed: {ex.Message}");
                return $"Connection failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Closes the connection cleanly: unsubscribes everything, stops the receive
        /// loop and resets state so a later ConnectAsync starts from scratch.
        /// </summary>
        public async Task<string> DisconnectAsync()
        {
            try
            {
                var socket = _webSocket;
                if (socket == null)
                    return "Not connected";

                if (socket.State == WebSocketState.Open)
                {
                    try { await UnsubscribeAllAsync(); } catch { }
                    if (_ordersSubscribed)
                    {
                        try { await UnsubscribeOrdersAsync(); } catch { }
                    }

                    try
                    {
                        using var closeCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                        await socket.CloseOutputAsync(WebSocketCloseStatus.NormalClosure, "Client disconnect", closeCts.Token);
                    }
                    catch (Exception ex)
                    {
                        Log("WARN", $"Close handshake failed: {ex.Message}");
                    }
                }

                try { _cancellationTokenSource?.Cancel(); } catch { }

                var receiveTask = _receiveTask;
                if (receiveTask != null)
                {
                    try { await Task.WhenAny(receiveTask, Task.Delay(2000)); } catch { }
                }

                try { socket.Dispose(); } catch { }
                try { _cancellationTokenSource?.Dispose(); } catch { }

                _webSocket = null;
                _cancellationTokenSource = null;
                _receiveTask = null;
                _isAuthenticated = false;
                Volatile.Write(ref _connectingFlag, 0);
                _ordersSubscribed = false;
                _orderSubscribeInProgress = false;

                // Disconnecting is not the same as the user asking to stop order updates.
                // UnsubscribeOrdersAsync above sets the manual flag, and only
                // SubscribeOrdersAsync clears it, so leaving it set would make
                // oa_ws_orders report "Unsubscribed" for the rest of the Excel session
                // even after a successful reconnect. UnsubscribeAllAsync clears the
                // market data equivalent for the same reason.
                _ordersManuallyUnsubscribed = false;

                _subscriptions.Clear();
                _marketData.Clear();
                _topicAliases.Clear();
                _autoSubscribeInProgress.Clear();
                CancelAllThrottles();
                _warnedUnsubscribed.Clear();
                _pendingSubscriptions.Clear();
                _pendingRequests.Clear();

                NotifyAllStatus("Disconnected");

                Log("INFO", "Disconnected from WebSocket server");
                return "Disconnected";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Disconnect failed: {ex.Message}");
                return $"Disconnect failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Authenticates with the WebSocket server
        /// </summary>
        private async Task<string> AuthenticateAsync()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey))
                {
                    Log("ERROR", "API Key not set");
                    return "Error: API Key not set. Use oa_api() first.";
                }

                var authMessage = new JObject
                {
                    ["action"] = "authenticate",
                    ["api_key"] = OpenAlgoConfig.ApiKey
                };

                Log("INFO", "Sending authentication request");
                await SendMessageAsync(authMessage.ToString());

                // Legacy mode: For backward compatibility with backends that don't send confirmations
                if (_legacyMode)
                {
                    // Wait briefly for authentication to process
                    await Task.Delay(500);
                    _isAuthenticated = true;
                    Log("INFO", "Authentication assumed successful (legacy mode)");
                    return "Connected and authenticated";
                }

                // Modern mode: Wait for actual server confirmation
                _authTcs = new TaskCompletionSource<bool>();

                // Wait for ACTUAL authentication response (max 30 seconds)
                var timeoutTask = Task.Delay(30000);
                var completedTask = await Task.WhenAny(_authTcs.Task, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    _isAuthenticated = false;
                    Log("ERROR", "Authentication timeout - no response from server after 30 seconds");
                    return "Error: Authentication timeout - no response from server";
                }

                bool authSuccess = await _authTcs.Task;
                _isAuthenticated = authSuccess;

                if (authSuccess)
                {
                    Log("INFO", "Authentication successful");
                    return "Connected and authenticated";
                }
                else
                {
                    Log("ERROR", "Authentication failed - server rejected credentials");
                    return "Error: Authentication failed - server rejected credentials";
                }
            }
            catch (Exception ex)
            {
                _isAuthenticated = false;
                Log("ERROR", $"Authentication exception: {ex.Message}");
                return $"Authentication failed: {ex.Message}";
            }
            finally
            {
                _authTcs = null;
            }
        }

        /// <summary>
        /// True when a ConnectAsync result means the socket is usable.
        ///
        /// ConnectAsync reports success as "Connected and authenticated" but also returns
        /// "Already connected" and "Connection already in progress" when another cell got
        /// there first. Those are healthy outcomes, not failures. A plain
        /// Contains("Connected") test is case sensitive and rejects both, which made every
        /// concurrent auto-subscribe on a sheet fail behind the one that won the race.
        /// </summary>
        private static bool IsConnectSuccess(string result)
        {
            if (string.IsNullOrWhiteSpace(result))
                return false;

            if (result.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                return false;

            return result.IndexOf("connected", StringComparison.OrdinalIgnoreCase) >= 0
                || result.IndexOf("already in progress", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>
        /// True when a subscribe result string reports success rather than a failure.
        /// </summary>
        private static bool SubscribeSucceeded(string result)
        {
            if (string.IsNullOrWhiteSpace(result))
                return false;

            return !result.StartsWith("Error", StringComparison.OrdinalIgnoreCase)
                && result.IndexOf("failed", StringComparison.OrdinalIgnoreCase) < 0
                && result.IndexOf("Not connected", StringComparison.OrdinalIgnoreCase) < 0
                && result.IndexOf("Not authenticated", StringComparison.OrdinalIgnoreCase) < 0;
        }

        /// <summary>
        /// Cancels and releases the current socket and its cancellation token source.
        ///
        /// Called before a reconnect so a previous dead attempt cannot leave an orphaned
        /// receive loop running against a socket nobody references any more. Safe to call
        /// when there is nothing to release.
        /// </summary>
        private void DisposeSocket()
        {
            var socket = _webSocket;
            var cts = _cancellationTokenSource;

            _webSocket = null;
            _cancellationTokenSource = null;

            try { cts?.Cancel(); } catch { }
            try { cts?.Dispose(); } catch { }
            try { socket?.Dispose(); } catch { }
        }

        /// <summary>
        /// Waits briefly for a connection started by another caller to become open and
        /// authenticated. Returns as soon as it is ready, or after the timeout.
        /// </summary>
        private async Task WaitForReadyAsync(int timeoutMs = 5000)
        {
            const int PollMs = 50;
            for (int waited = 0; waited < timeoutMs; waited += PollMs)
            {
                if (_webSocket?.State == WebSocketState.Open && _isAuthenticated)
                    return;

                // A connection attempt that has finished and failed will not recover by
                // waiting, so stop as soon as nobody is still trying.
                if (!_isConnecting && _webSocket?.State != WebSocketState.Open)
                    return;

                await Task.Delay(PollMs).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Makes sure the socket is open and authenticated, connecting on demand.
        /// Returns null when ready, otherwise the message to show in the cell.
        /// </summary>
        private async Task<string?> EnsureReadyAsync(string context)
        {
            if (_webSocket?.State != WebSocketState.Open)
            {
                var connectResult = await ConnectAsync();
                if (!IsConnectSuccess(connectResult))
                    return connectResult;

                // Another cell may still be connecting. Opening a sheet creates many
                // streaming cells at once and they all reach this point together, so
                // wait for the winner to finish rather than subscribing on a socket
                // that is not open yet.
                await WaitForReadyAsync();
            }

            if (_isAuthenticated)
                return null;

            // Try to re-authenticate if API key is now available.
            // This handles the case where a previous attempt left WS connected but not authed.
            if (!string.IsNullOrWhiteSpace(OpenAlgoConfig.ApiKey) && _webSocket?.State == WebSocketState.Open)
            {
                Log("INFO", $"Attempting re-authentication before {context}");
                var authResult = await AuthenticateAsync();
                if (!_isAuthenticated)
                {
                    Log("ERROR", $"Re-authentication failed for {context}: {authResult}");
                    return $"Error: {authResult}";
                }
                return null;
            }

            Log("ERROR", $"Cannot {context} - not authenticated");
            return "Error: Not authenticated. Use oa_api() first.";
        }

        // ------------------------------------------------------------------
        // Market data subscriptions
        // ------------------------------------------------------------------

        /// <summary>
        /// Subscribes to market data for a symbol
        /// </summary>
        public async Task<string> SubscribeAsync(string symbol, string exchange, int mode, int? depthLevel = null)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);

            try
            {
                var notReady = await EnsureReadyAsync($"subscribe to {key}");
                if (notReady != null)
                    return notReady;

                var subscribeMessage = new JObject
                {
                    ["action"] = "subscribe",
                    ["symbol"] = symbol,
                    ["exchange"] = exchange,
                    ["mode"] = mode
                };

                if (mode == 3 && depthLevel.HasValue)
                {
                    subscribeMessage["depth_level"] = depthLevel.Value;
                    _requestedDepth[key] = depthLevel.Value;
                }

                Log("INFO", $"Subscribing to {key}");
                await SendMessageAsync(subscribeMessage.ToString());

                // Legacy mode: For backward compatibility with backends that don't send confirmations
                if (_legacyMode)
                {
                    // Immediately add to subscriptions (assume success)
                    _subscriptions[key] = new SubscriptionInfo
                    {
                        Symbol = symbol,
                        Exchange = exchange,
                        Mode = mode,
                        DepthLevel = depthLevel,
                        SubscribedAt = DateTime.UtcNow
                    };

                    // Clear manual unsubscribe flag to allow this subscription
                    _manuallyUnsubscribed.TryRemove(key, out _);
                    _warnedUnsubscribed.TryRemove(key, out _);

                    Log("INFO", $"Subscribed to {key} (legacy mode - assumed success)");
                    NotifyStatus(key, "Waiting for data...");
                    return $"Subscribed: {symbol} ({exchange}) - Mode {mode}";
                }

                // Modern mode: Wait for server confirmation
                var tcs = new TaskCompletionSource<bool>();
                _pendingSubscriptions[key] = tcs;

                // Wait for server confirmation (max 30 seconds)
                var timeoutTask = Task.Delay(30000);
                var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    _pendingSubscriptions.TryRemove(key, out _);
                    Log("ERROR", $"Subscription timeout for {key} after 30 seconds");
                    return $"Error: Subscription timeout for {symbol}";
                }

                bool subSuccess = await tcs.Task;

                if (subSuccess)
                {
                    // Only add to subscriptions if server confirmed
                    _subscriptions[key] = new SubscriptionInfo
                    {
                        Symbol = symbol,
                        Exchange = exchange,
                        Mode = mode,
                        DepthLevel = depthLevel,
                        SubscribedAt = DateTime.UtcNow
                    };

                    // Clear manual unsubscribe flag to allow this subscription
                    _manuallyUnsubscribed.TryRemove(key, out _);
                    _warnedUnsubscribed.TryRemove(key, out _);

                    Log("INFO", $"Successfully subscribed to {key}");
                    NotifyStatus(key, "Waiting for data...");
                    return $"Subscribed: {symbol} ({exchange}) - Mode {mode}";
                }
                else
                {
                    Log("ERROR", $"Server rejected subscription for {key}");
                    return $"Error: Server rejected subscription for {symbol}";
                }
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Subscribe exception for {key}: {ex.Message}");
                return $"Subscribe failed: {ex.Message}";
            }
            finally
            {
                _pendingSubscriptions.TryRemove(key, out _);
            }
        }

        /// <summary>
        /// Unsubscribes from market data
        /// </summary>
        public async Task<string> UnsubscribeAsync(string symbol, string exchange, int mode)
        {
            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                    return "Error: Not connected";

                var unsubscribeMessage = new JObject
                {
                    ["action"] = "unsubscribe",
                    ["symbol"] = symbol,
                    ["exchange"] = exchange,
                    ["mode"] = mode
                };

                await SendMessageAsync(unsubscribeMessage.ToString());

                // Remove subscription info and mark as manually unsubscribed
                string key = GetSubscriptionKey(symbol, exchange, mode);
                _subscriptions.TryRemove(key, out _);

                // Remove from market data cache. The previous version rebuilt
                // "symbol|exchange|mode" here, which is exactly what GetSubscriptionKey
                // already produced, so it removed the same entry twice and never evicted
                // the entry stored under the server's own topic string. That alias is now
                // tracked so it can actually be released.
                _marketData.TryRemove(key, out _);

                if (_topicAliases.TryRemove(key, out var alias))
                    _marketData.TryRemove(alias, out _);

                // Mark this as manually unsubscribed to prevent auto-resubscribe
                _manuallyUnsubscribed[key] = DateTime.UtcNow;

                NotifyStatus(key, "Unsubscribed");

                Log("INFO", $"Unsubscribed from {key}");
                return $"Unsubscribed: {symbol} ({exchange}) - Mode {mode}";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Unsubscribe failed for {symbol}: {ex.Message}");
                return $"Unsubscribe failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Gets the latest market data for a subscription
        /// </summary>
        public JObject? GetMarketData(string symbol, string exchange, int mode)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);
            return _marketData.TryGetValue(key, out var data) ? data : null;
        }

        /// <summary>
        /// Gets all active subscriptions
        /// </summary>
        public string[] GetActiveSubscriptions()
        {
            return _subscriptions.Keys.ToArray();
        }

        /// <summary>
        /// Checks if a subscription exists
        /// </summary>
        public bool IsSubscribed(string symbol, string exchange, int mode)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);
            return _subscriptions.ContainsKey(key);
        }

        /// <summary>
        /// Starts a non-blocking auto-subscribe in the background.
        /// Prevents duplicate attempts for the same subscription key.
        /// Used by the streaming Excel functions (oa_ws_ltp, oa_ws_quote, etc.) to avoid
        /// blocking the calculation thread, which would deadlock with ConnectAsync's
        /// API key wait loop.
        /// </summary>
        public void StartAutoSubscribe(string symbol, string exchange, int mode, int? depthLevel = null)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);
            if (depthLevel.HasValue)
                _requestedDepth[key] = depthLevel.Value;

            if (!_autoSubscribeInProgress.TryAdd(key, true))
                return; // Already in progress

            _ = Task.Run(async () =>
            {
                try
                {
                    string result = await SubscribeAsync(symbol, exchange, mode, depthLevel);

                    // A failure that comes back as a returned string rather than an
                    // exception must still reach the cell. The RTD topic is created once
                    // and nothing re-runs this, so swallowing the message would leave the
                    // cell on "Subscribing..." forever with the reason only in the log.
                    if (!SubscribeSucceeded(result))
                    {
                        Log("ERROR", $"Auto-subscribe failed for {key}: {result}");
                        NotifyStatus(key, result);
                    }
                }
                catch (Exception ex)
                {
                    Log("ERROR", $"Auto-subscribe failed for {key}: {ex.Message}");
                    NotifyStatus(key, $"Error: {ex.Message}");
                }
                finally
                {
                    _autoSubscribeInProgress.TryRemove(key, out _);
                }
            });
        }

        /// <summary>
        /// Checks if a symbol was manually unsubscribed (to prevent auto-resubscribe)
        /// </summary>
        public bool WasManuallyUnsubscribed(string symbol, string exchange, int mode)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);
            return _manuallyUnsubscribed.ContainsKey(key);
        }

        /// <summary>
        /// Clears the manual unsubscribe flag (called during disconnect/reconnect)
        /// </summary>
        public void ClearManualUnsubscribeFlags()
        {
            _manuallyUnsubscribed.Clear();
        }

        /// <summary>
        /// Unsubscribes from all active subscriptions
        /// </summary>
        public async Task<string> UnsubscribeAllAsync()
        {
            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                    return "Not connected";

                int count = _subscriptions.Count;
                if (count == 0)
                    return "No active subscriptions";

                // Create a copy of subscription keys to avoid modification during iteration
                var subscriptionKeys = _subscriptions.Keys.ToArray();

                foreach (var key in subscriptionKeys)
                {
                    if (_subscriptions.TryGetValue(key, out var subInfo))
                    {
                        await UnsubscribeAsync(subInfo.Symbol, subInfo.Exchange, subInfo.Mode);
                    }
                }

                // Clear manual unsubscribe flags after unsubscribe_all
                // This allows symbols to be resubscribed later (unsubscribe_all is a "reset")
                ClearManualUnsubscribeFlags();

                // Clear all market data cache to stop data from showing in cells
                _marketData.Clear();
                _topicAliases.Clear();

                Log("INFO", $"Unsubscribed from all {count} subscription(s), cleared manual unsubscribe flags, and cleared market data cache");
                return $"Unsubscribed from {count} subscription(s)";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Unsubscribe all failed: {ex.Message}");
                return $"Unsubscribe all failed: {ex.Message}";
            }
        }

        // ------------------------------------------------------------------
        // Order updates
        // ------------------------------------------------------------------

        /// <summary>
        /// Subscribes to the account wide order update stream (action subscribe_orders).
        /// </summary>
        public async Task<string> SubscribeOrdersAsync()
        {
            try
            {
                var notReady = await EnsureReadyAsync("subscribe to order updates");
                if (notReady != null)
                    return notReady;

                var message = new JObject
                {
                    ["action"] = "subscribe_orders",
                    ["request_id"] = "excel-orders"
                };

                await SendMessageAsync(message.ToString());
                _ordersSubscribed = true;
                _ordersManuallyUnsubscribed = false;

                Log("INFO", "Subscribed to order updates");
                NotifyStatus(OrdersTopic, "Waiting for order updates...");
                return "Subscribed to order updates";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Subscribe orders failed: {ex.Message}");
                return $"Subscribe orders failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Unsubscribes from the order update stream (action unsubscribe_orders).
        /// </summary>
        public async Task<string> UnsubscribeOrdersAsync()
        {
            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                    return "Error: Not connected";

                var message = new JObject
                {
                    ["action"] = "unsubscribe_orders",
                    ["request_id"] = "excel-orders"
                };

                await SendMessageAsync(message.ToString());
                _ordersSubscribed = false;
                _ordersManuallyUnsubscribed = true;

                NotifyStatus(OrdersTopic, "Unsubscribed");

                Log("INFO", "Unsubscribed from order updates");
                return "Unsubscribed from order updates";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Unsubscribe orders failed: {ex.Message}");
                return $"Unsubscribe orders failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Starts a non-blocking subscribe to order updates, mirroring StartAutoSubscribe.
        /// </summary>
        public void StartAutoSubscribeOrders()
        {
            if (_orderSubscribeInProgress)
                return;

            _orderSubscribeInProgress = true;
            _ = Task.Run(async () =>
            {
                try
                {
                    string result = await SubscribeOrdersAsync();
                    if (!SubscribeSucceeded(result))
                    {
                        Log("ERROR", $"Auto-subscribe to order updates failed: {result}");
                        NotifyStatus(OrdersTopic, result);
                    }
                }
                catch (Exception ex)
                {
                    Log("ERROR", $"Auto-subscribe to order updates failed: {ex.Message}");
                    NotifyStatus(OrdersTopic, $"Error: {ex.Message}");
                }
                finally
                {
                    _orderSubscribeInProgress = false;
                }
            });
        }

        /// <summary>
        /// True once subscribe_orders has been sent and not withdrawn.
        /// </summary>
        public bool IsSubscribedToOrders => _ordersSubscribed;

        /// <summary>
        /// True when the user explicitly called unsubscribe_orders.
        /// </summary>
        public bool OrdersManuallyUnsubscribed => _ordersManuallyUnsubscribed;

        /// <summary>
        /// Returns a snapshot of the buffered order updates, newest first.
        /// </summary>
        public JObject[] GetOrderUpdates()
        {
            lock (_orderLock)
            {
                return _orderUpdates.ToArray();
            }
        }

        private void StoreOrderUpdate(JObject update)
        {
            lock (_orderLock)
            {
                _orderUpdates.Insert(0, update);
                while (_orderUpdates.Count > MaxOrderUpdates)
                    _orderUpdates.RemoveAt(_orderUpdates.Count - 1);
            }
        }

        // ------------------------------------------------------------------
        // Request/response actions
        // ------------------------------------------------------------------

        /// <summary>
        /// Sends an action and waits for the matching response message.
        /// Returns the response, or an object carrying status "error" on failure.
        /// </summary>
        public async Task<JObject> SendRequestAsync(string action, string responseType, JObject? extra = null, int timeoutMs = 15000)
        {
            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                {
                    var connectResult = await ConnectAsync();
                    if (!IsConnectSuccess(connectResult))
                        return ErrorObject(connectResult);
                }

                var tcs = new TaskCompletionSource<JObject>(TaskCreationOptions.RunContinuationsAsynchronously);
                _pendingRequests[responseType] = tcs;

                try
                {
                    var request = new JObject { ["action"] = action };
                    if (extra != null)
                    {
                        foreach (var property in extra.Properties())
                            request[property.Name] = property.Value;
                    }

                    await SendMessageAsync(request.ToString());

                    var completed = await Task.WhenAny(tcs.Task, Task.Delay(timeoutMs));
                    if (completed != tcs.Task)
                        return ErrorObject($"Timed out waiting for {responseType}");

                    return await tcs.Task;
                }
                finally
                {
                    _pendingRequests.TryRemove(responseType, out _);
                }
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Request {action} failed: {ex.Message}");
                return ErrorObject(ex.Message);
            }
        }

        /// <summary>
        /// Asks the server for the connected broker (action get_broker_info).
        /// </summary>
        public Task<JObject> GetBrokerInfoAsync() =>
            SendRequestAsync("get_broker_info", "broker_info");

        /// <summary>
        /// Asks the server for the list of supported brokers (action get_supported_brokers).
        /// </summary>
        public Task<JObject> GetSupportedBrokersAsync() =>
            SendRequestAsync("get_supported_brokers", "supported_brokers");

        /// <summary>
        /// Sends a ping and reports the round trip in milliseconds.
        /// </summary>
        public async Task<JObject> PingAsync()
        {
            var stopwatch = Stopwatch.StartNew();
            var extra = new JObject
            {
                ["timestamp"] = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds()
            };

            var response = await SendRequestAsync("ping", "pong", extra);
            stopwatch.Stop();

            if (response["status"]?.ToString() != "error")
                response["round_trip_ms"] = stopwatch.ElapsedMilliseconds;

            return response;
        }

        private static JObject ErrorObject(string message) =>
            new JObject { ["status"] = "error", ["message"] = message };

        // ------------------------------------------------------------------
        // Transport
        // ------------------------------------------------------------------

        /// <summary>
        /// Sends a message through the WebSocket
        /// </summary>
        private async Task SendMessageAsync(string message)
        {
            if (_webSocket?.State != WebSocketState.Open)
                throw new InvalidOperationException("WebSocket is not connected");

            byte[] buffer = Encoding.UTF8.GetBytes(message);
            await _webSocket.SendAsync(new ArraySegment<byte>(buffer), WebSocketMessageType.Text, true, CancellationToken.None);
        }

        /// <summary>
        /// Continuously receives messages from the WebSocket
        /// </summary>
        private async Task ReceiveMessagesAsync(CancellationToken cancellationToken)
        {
            var buffer = new byte[1024 * 16]; // 16KB buffer
            var messageBuilder = new StringBuilder();

            try
            {
                Log("INFO", "Started receiving WebSocket messages");

                while (!cancellationToken.IsCancellationRequested && _webSocket?.State == WebSocketState.Open)
                {
                    var result = await _webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        Log("INFO", "WebSocket close message received");
                        await _webSocket.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closing", CancellationToken.None);
                        break;
                    }

                    messageBuilder.Append(Encoding.UTF8.GetString(buffer, 0, result.Count));

                    if (result.EndOfMessage)
                    {
                        string message = messageBuilder.ToString();
                        messageBuilder.Clear();

                        // Process the message
                        await ProcessMessageAsync(message);
                    }
                }

                Log("INFO", "Stopped receiving WebSocket messages");
            }
            catch (OperationCanceledException)
            {
                Log("INFO", "WebSocket receive task cancelled");
            }
            catch (Exception ex)
            {
                Log("ERROR", $"WebSocket receive error: {ex.Message}");
            }
        }

        /// <summary>
        /// Processes incoming WebSocket messages
        /// </summary>
        private async Task ProcessMessageAsync(string message)
        {
            try
            {
                // Handle ping messages
                if (message.Trim().Equals("ping", StringComparison.OrdinalIgnoreCase))
                {
                    await SendMessageAsync("pong");
                    _lastPingTime = DateTime.UtcNow;
                    return;
                }

                // Parse JSON message
                var jsonMessage = JObject.Parse(message);

                // Check message type
                string? messageType = jsonMessage["type"]?.ToString();

                if (messageType == "market_data")
                {
                    HandleMarketData(jsonMessage);
                }
                else if (messageType == "authentication" || messageType == "auth")
                {
                    string? status = jsonMessage["status"]?.ToString();
                    bool success = status == "success";
                    _isAuthenticated = success;

                    // Complete the authentication task
                    _authTcs?.TrySetResult(success);

                    Log("INFO", $"Authentication response: {status}");
                }
                else if (messageType == "subscription")
                {
                    // Handle subscription confirmation
                    string? status = jsonMessage["status"]?.ToString();
                    string? symbol = jsonMessage["symbol"]?.ToString();
                    string? exchange = jsonMessage["exchange"]?.ToString();
                    int? mode = jsonMessage["mode"]?.ToObject<int?>();

                    if (symbol != null && exchange != null && mode.HasValue)
                    {
                        string key = GetSubscriptionKey(symbol, exchange, mode.Value);

                        if (_pendingSubscriptions.TryGetValue(key, out var tcs))
                        {
                            bool success = status == "success";
                            tcs.TrySetResult(success);

                            Log("INFO", $"Subscription confirmation for {key}: {status}");
                        }
                    }
                }
                else if (messageType == "order_update")
                {
                    StoreOrderUpdate(jsonMessage);
                    NotifyTopic(OrdersTopic, jsonMessage, bypassThrottle: true);
                    Log("INFO", $"Order update: {jsonMessage["orderid"]} {jsonMessage["order_status"]}");
                }
                else if (messageType == "subscribe_orders")
                {
                    _ordersSubscribed = jsonMessage["status"]?.ToString() == "success";
                    Log("INFO", $"subscribe_orders response: {jsonMessage["status"]}");
                    CompletePendingRequest("subscribe_orders", jsonMessage);
                }
                else if (messageType == "unsubscribe_orders")
                {
                    _ordersSubscribed = false;
                    Log("INFO", $"unsubscribe_orders response: {jsonMessage["status"]}");
                    CompletePendingRequest("unsubscribe_orders", jsonMessage);
                }
                else if (messageType == "broker_info"
                         || messageType == "supported_brokers"
                         || messageType == "pong")
                {
                    CompletePendingRequest(messageType, jsonMessage);
                }
                else if (messageType == null && jsonMessage["status"]?.ToString() == "error")
                {
                    // The server reports errors without a type field, so the only safe
                    // correlation is to fail whatever request is currently outstanding.
                    string errorText = jsonMessage["message"]?.ToString() ?? "Unknown error";
                    Log("ERROR", $"Server error: {jsonMessage["code"]} {errorText}");

                    foreach (var pending in _pendingRequests.Keys.ToArray())
                        CompletePendingRequest(pending, jsonMessage);
                }
                else
                {
                    Log("DEBUG", $"Received unknown message type: {messageType}");
                }
            }
            catch (JsonException ex)
            {
                Log("ERROR", $"Invalid JSON message: {ex.Message}");
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Error processing message: {ex.Message}");
            }
        }

        private void CompletePendingRequest(string responseType, JObject response)
        {
            if (_pendingRequests.TryRemove(responseType, out var tcs))
                tcs.TrySetResult(response);
        }

        /// <summary>
        /// Stores an incoming tick and wakes the cells watching that topic.
        /// </summary>
        private void HandleMarketData(JObject jsonMessage)
        {
            // Extract subscription details - support both formats
            // Format 1: symbol/exchange inside data object
            // Format 2: symbol/exchange at top level (OpenAlgo backend format)
            var data = jsonMessage["data"] as JObject;
            string? symbol = data?["symbol"]?.ToString() ?? jsonMessage["symbol"]?.ToString();
            string? exchange = data?["exchange"]?.ToString() ?? jsonMessage["exchange"]?.ToString();
            int? mode = jsonMessage["mode"]?.ToObject<int?>();

            if (string.IsNullOrEmpty(symbol) || string.IsNullOrEmpty(exchange) || !mode.HasValue)
            {
                Log("WARN", $"Invalid market_data message - symbol: {symbol}, exchange: {exchange}, mode: {mode}");
                return;
            }

            string key = GetSubscriptionKey(symbol, exchange, mode.Value);

            // Only store data if we have an active subscription
            if (!_subscriptions.ContainsKey(key))
            {
                // Log once per topic: this can otherwise fire on every tick.
                if (_warnedUnsubscribed.TryAdd(key, true))
                    Log("WARN", $"Received data for unsubscribed symbol: {key}");
                return;
            }

            // If symbol/exchange are at top level, move them into data for consistency
            if (jsonMessage["symbol"] != null && data != null)
            {
                if (data["symbol"] == null) data["symbol"] = symbol;
                if (data["exchange"] == null) data["exchange"] = exchange;
            }

            // The cache always takes every tick; only the push to Excel is throttled.
            _marketData[key] = jsonMessage;

            // Also store with the topic key for backward compatibility
            string? topic = jsonMessage["topic"]?.ToString();
            if (topic != null && topic != key)
            {
                _marketData[topic] = jsonMessage;
                _topicAliases[key] = topic;
            }

            NotifyTopic(key, jsonMessage);
        }

        /// <summary>
        /// Generates a unique key for subscription tracking
        /// </summary>
        private string GetSubscriptionKey(string symbol, string exchange, int mode)
        {
            return TopicFor(symbol, exchange, mode);
        }
    }

    /// <summary>
    /// Information about an active subscription
    /// </summary>
    public class SubscriptionInfo
    {
        public string Symbol { get; set; } = "";
        public string Exchange { get; set; } = "";
        public int Mode { get; set; }
        public int? DepthLevel { get; set; }
        public DateTime SubscribedAt { get; set; }
    }
}

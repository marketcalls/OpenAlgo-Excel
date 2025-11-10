using System;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using ExcelDna.Integration;
using System.Timers;
using System.IO;

namespace OpenAlgo
{
    /// <summary>
    /// Manages WebSocket connections for real-time market data streaming from OpenAlgo
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

        private bool _isAuthenticated = false;
        private bool _isConnecting = false;
        private DateTime _lastPingTime = DateTime.MinValue;

        // WebSocket configuration
        private string _wsUrl = "ws://127.0.0.1:8765";

        // Timer for triggering Excel updates (continuous streaming)
        private System.Timers.Timer? _updateTimer;

        // Authentication confirmation tracking
        private TaskCompletionSource<bool>? _authTcs;

        // Subscription confirmation tracking
        private readonly ConcurrentDictionary<string, TaskCompletionSource<bool>> _pendingSubscriptions = new();

        // Performance optimization - debounce Excel updates
        private DateTime _lastExcelUpdate = DateTime.MinValue;
        private const int MIN_UPDATE_INTERVAL_MS = 50; // Max 20 updates/second

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
            // Initialize update timer - triggers Excel recalculation continuously for live streaming
            _updateTimer = new System.Timers.Timer(100); // 100ms = 10 updates per second (balanced performance)
            _updateTimer.Elapsed += OnUpdateTimerElapsed;
            _updateTimer.AutoReset = true;
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
        /// Timer event handler to trigger Excel recalculation
        /// </summary>
        private void OnUpdateTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            // Continuously trigger Excel recalculation for live streaming
            // This ensures cells update in real-time as WebSocket data arrives
            if (_webSocket?.State == WebSocketState.Open && _isAuthenticated)
            {
                TriggerExcelUpdate();
            }
        }

        /// <summary>
        /// Triggers Excel recalculation to update all volatile functions (with debouncing)
        /// </summary>
        private void TriggerExcelUpdate()
        {
            try
            {
                // Debounce: only update if enough time has passed
                var now = DateTime.UtcNow;
                if ((now - _lastExcelUpdate).TotalMilliseconds < MIN_UPDATE_INTERVAL_MS)
                    return;

                _lastExcelUpdate = now;

                // Force Excel to recalculate volatile functions using COM automation
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        var app = ExcelDnaUtil.Application as dynamic;
                        if (app != null)
                        {
                            // Calculate all volatile functions to ensure continuous updates
                            app.Calculate();
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("ERROR", $"Excel calculation failed: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                Log("ERROR", $"TriggerExcelUpdate failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the WebSocket URL for connection
        /// </summary>
        public void SetWebSocketUrl(string url)
        {
            _wsUrl = url;
        }

        /// <summary>
        /// Gets the current connection state
        /// </summary>
        public string GetConnectionState()
        {
            if (_webSocket == null) return "Disconnected";
            return _webSocket.State.ToString();
        }

        /// <summary>
        /// Connects to the WebSocket server and authenticates
        /// </summary>
        public async Task<string> ConnectAsync()
        {
            try
            {
                if (_isConnecting)
                {
                    Log("WARN", "Connection already in progress");
                    return "Connection already in progress";
                }

                if (_webSocket?.State == WebSocketState.Open)
                {
                    Log("INFO", "Already connected");
                    return "Already connected";
                }

                _isConnecting = true;
                Log("INFO", $"Connecting to {_wsUrl}");

                _webSocket = new ClientWebSocket();
                _cancellationTokenSource = new CancellationTokenSource();

                await _webSocket.ConnectAsync(new Uri(_wsUrl), _cancellationTokenSource.Token);
                Log("INFO", "WebSocket connected successfully");

                // Start receiving messages
                _receiveTask = Task.Run(() => ReceiveMessagesAsync(_cancellationTokenSource.Token));

                // Authenticate
                var authResult = await AuthenticateAsync();

                // Start the update timer for continuous streaming
                _updateTimer?.Start();

                _isConnecting = false;
                return authResult;
            }
            catch (Exception ex)
            {
                _isConnecting = false;
                Log("ERROR", $"Connection failed: {ex.Message}");
                return $"Connection failed: {ex.Message}";
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
        /// Subscribes to market data for a symbol
        /// </summary>
        public async Task<string> SubscribeAsync(string symbol, string exchange, int mode, int? depthLevel = null)
        {
            string key = GetSubscriptionKey(symbol, exchange, mode);

            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                {
                    var connectResult = await ConnectAsync();
                    if (!connectResult.Contains("Connected"))
                        return connectResult;
                }

                if (!_isAuthenticated)
                {
                    Log("ERROR", $"Cannot subscribe to {key} - not authenticated");
                    return "Error: Not authenticated";
                }

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

                    Log("INFO", $"Subscribed to {key} (legacy mode - assumed success)");
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

                    Log("INFO", $"Successfully subscribed to {key}");
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

                // Remove from market data cache - BOTH keys to prevent memory leak
                _marketData.TryRemove(key, out _);

                // Also remove topic-based key (backward compatibility storage)
                string topic = $"{symbol}|{exchange}|{mode}";
                _marketData.TryRemove(topic, out _);

                // Mark this as manually unsubscribed to prevent auto-resubscribe
                _manuallyUnsubscribed[key] = DateTime.UtcNow;

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

                Log("INFO", $"Unsubscribed from all {count} subscription(s), cleared manual unsubscribe flags, and cleared market data cache");
                return $"Unsubscribed from {count} subscription(s)";
            }
            catch (Exception ex)
            {
                Log("ERROR", $"Unsubscribe all failed: {ex.Message}");
                return $"Unsubscribe all failed: {ex.Message}";
            }
        }

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
                    // Extract subscription details from the data
                    var data = jsonMessage["data"] as JObject;
                    string? symbol = data?["symbol"]?.ToString();
                    string? exchange = data?["exchange"]?.ToString();
                    int? mode = jsonMessage["mode"]?.ToObject<int?>();

                    // Only store data if we have an active subscription
                    if (!string.IsNullOrEmpty(symbol) && !string.IsNullOrEmpty(exchange) && mode.HasValue)
                    {
                        string key = GetSubscriptionKey(symbol, exchange, mode.Value);

                        // Check if this subscription is still active
                        if (_subscriptions.ContainsKey(key))
                        {
                            _marketData[key] = jsonMessage;

                            // Also store with the topic key for backward compatibility
                            string? topic = jsonMessage["topic"]?.ToString();
                            if (topic != null)
                            {
                                _marketData[topic] = jsonMessage;
                            }

                            // Immediately trigger Excel update for responsive streaming
                            TriggerExcelUpdate();
                        }
                        else
                        {
                            Log("WARN", $"Received data for unsubscribed symbol: {key}");
                        }
                    }
                }
                else if (messageType == "authentication")
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

        /// <summary>
        /// Generates a unique key for subscription tracking
        /// </summary>
        private string GetSubscriptionKey(string symbol, string exchange, int mode)
        {
            return $"{symbol}|{exchange}|{mode}";
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

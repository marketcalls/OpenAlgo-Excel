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
        /// Triggers Excel recalculation to update all volatile functions
        /// </summary>
        private void TriggerExcelUpdate()
        {
            try
            {
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
                    catch
                    {
                        // Ignore if calculation fails
                    }
                });
            }
            catch
            {
                // Ignore if Excel interface not available
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
                    return "Connection already in progress";

                if (_webSocket?.State == WebSocketState.Open)
                    return "Already connected";

                _isConnecting = true;

                _webSocket = new ClientWebSocket();
                _cancellationTokenSource = new CancellationTokenSource();

                await _webSocket.ConnectAsync(new Uri(_wsUrl), _cancellationTokenSource.Token);

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
                    return "Error: API Key not set. Use oa_api() first.";

                var authMessage = new JObject
                {
                    ["action"] = "authenticate",
                    ["api_key"] = OpenAlgoConfig.ApiKey
                };

                await SendMessageAsync(authMessage.ToString());

                // Wait a bit for authentication response
                await Task.Delay(500);

                _isAuthenticated = true;
                return "Connected and authenticated";
            }
            catch (Exception ex)
            {
                return $"Authentication failed: {ex.Message}";
            }
        }

        /// <summary>
        /// Subscribes to market data for a symbol
        /// </summary>
        public async Task<string> SubscribeAsync(string symbol, string exchange, int mode, int? depthLevel = null)
        {
            try
            {
                if (_webSocket?.State != WebSocketState.Open)
                {
                    var connectResult = await ConnectAsync();
                    if (!connectResult.Contains("Connected"))
                        return connectResult;
                }

                if (!_isAuthenticated)
                    return "Error: Not authenticated";

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

                await SendMessageAsync(subscribeMessage.ToString());

                // Store subscription info and clear manual unsubscribe flag
                string key = GetSubscriptionKey(symbol, exchange, mode);
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

                return $"Subscribed: {symbol} ({exchange}) - Mode {mode}";
            }
            catch (Exception ex)
            {
                return $"Subscribe failed: {ex.Message}";
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
                _marketData.TryRemove(key, out _);

                // Mark this as manually unsubscribed to prevent auto-resubscribe
                _manuallyUnsubscribed[key] = DateTime.UtcNow;

                return $"Unsubscribed: {symbol} ({exchange}) - Mode {mode}";
            }
            catch (Exception ex)
            {
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

                return $"Unsubscribed from {count} subscription(s)";
            }
            catch (Exception ex)
            {
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
                while (!cancellationToken.IsCancellationRequested && _webSocket?.State == WebSocketState.Open)
                {
                    var result = await _webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
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
            }
            catch (OperationCanceledException)
            {
                // Normal cancellation, ignore
            }
            catch (Exception ex)
            {
                Console.WriteLine($"WebSocket receive error: {ex.Message}");
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
                    }
                }
                else if (messageType == "authentication")
                {
                    string? status = jsonMessage["status"]?.ToString();
                    _isAuthenticated = status == "success";
                }
                else if (messageType == "subscription")
                {
                    // Handle subscription confirmation
                    string? status = jsonMessage["status"]?.ToString();
                    // Could log or handle subscription status here
                }
            }
            catch (JsonException)
            {
                // Invalid JSON, ignore
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing message: {ex.Message}");
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

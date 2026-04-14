using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ImapCertWatcher.Server
{
    public class TcpCommandServer
    {
        private readonly int _port;
        private readonly Func<string, string> _commandHandler;
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);
        private const int IoTimeoutMs = 60000;
        private const int MaxLoggedCommandLength = 300;

        private TcpListener _listener;
        private CancellationTokenSource _cts;

        public TcpCommandServer(int port, Func<string, string> commandHandler)
        {
            if (port <= 0 || port > 65535)
                throw new ArgumentOutOfRangeException(nameof(port), "Некорректный TCP-порт.");

            _commandHandler = commandHandler ?? throw new ArgumentNullException(nameof(commandHandler));
            _port = port;
        }

        public void Start()
        {
            if (_listener != null)
                return;

            _cts = new CancellationTokenSource();

            _listener = new TcpListener(IPAddress.Any, _port);

            _listener.Server.SetSocketOption(
                SocketOptionLevel.Socket,
                SocketOptionName.ReuseAddress,
                true);

            _listener.Start();

            Task.Run(AcceptLoop);
        }

        public void Stop()
        {
            var cts = _cts;
            var listener = _listener;

            _cts = null;
            _listener = null;

            try
            {
                cts?.Cancel();
            }
            catch
            {
            }

            try
            {
                listener?.Stop();
            }
            catch
            {
            }

            try
            {
                cts?.Dispose();
            }
            catch
            {
            }
        }

        private async Task AcceptLoop()
        {
            var listener = _listener;
            var cts = _cts;

            if (listener == null || cts == null)
                return;

            while (!cts.IsCancellationRequested)
            {
                try
                {
                    var client = await listener.AcceptTcpClientAsync();
                    _ = HandleClient(client);
                }
                catch (ObjectDisposedException)
                {
                    break;
                }
                catch (InvalidOperationException)
                {
                    break;
                }
                catch
                {
                    if (cts.IsCancellationRequested)
                        break;
                }
            }
        }

        private async Task HandleClient(TcpClient client)
        {
            using (client)
            using (var stream = client.GetStream())
            using (var reader = new StreamReader(stream, Utf8NoBom, false, 4096, true))
            using (var writer = new StreamWriter(stream, Utf8NoBom, 4096, true)
            {
                AutoFlush = true,
                NewLine = "\n"
            })
            {
                try
                {
                    client.ReceiveTimeout = IoTimeoutMs;
                    client.SendTimeout = IoTimeoutMs;

                    var request = await ReadLineWithTimeoutAsync(reader, IoTimeoutMs);

                    if (string.IsNullOrWhiteSpace(request))
                        return;

                    Console.WriteLine("[SERVER] CMD: " + DescribeCommand(request));

                    string response;
                    try
                    {
                        response = _commandHandler(request) ?? string.Empty;
                    }
                    catch (Exception ex)
                    {
                        response = "ERROR|SERVER|" + SanitizeSingleLine(ex.Message);
                        Console.WriteLine("[SERVER HANDLER ERROR] " + ex.Message);
                    }

                    await WriteLineWithTimeoutAsync(
                        writer,
                        response.Length.ToString(CultureInfo.InvariantCulture),
                        IoTimeoutMs);

                    if (response.Length > 0)
                        await WriteWithTimeoutAsync(writer, response, IoTimeoutMs);

                    await writer.FlushAsync();

                    Console.WriteLine("[SERVER] RESPONSE SENT");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[SERVER ERROR] " + ex.Message);
                }
            }
        }

        private static async Task<string> ReadLineWithTimeoutAsync(StreamReader reader, int timeoutMs)
        {
            var readTask = reader.ReadLineAsync();
            var completed = await Task.WhenAny(readTask, Task.Delay(timeoutMs));

            if (completed != readTask)
                throw new TimeoutException("Timeout while reading client request.");

            return await readTask;
        }

        private static async Task WriteLineWithTimeoutAsync(StreamWriter writer, string text, int timeoutMs)
        {
            var writeTask = writer.WriteLineAsync(text);
            var completed = await Task.WhenAny(writeTask, Task.Delay(timeoutMs));

            if (completed != writeTask)
                throw new TimeoutException("Timeout while writing response length.");

            await writeTask;
        }

        private static async Task WriteWithTimeoutAsync(StreamWriter writer, string text, int timeoutMs)
        {
            var writeTask = writer.WriteAsync(text);
            var completed = await Task.WhenAny(writeTask, Task.Delay(timeoutMs));

            if (completed != writeTask)
                throw new TimeoutException("Timeout while writing response body.");

            await writeTask;
        }

        private static string DescribeCommand(string request)
        {
            if (string.IsNullOrWhiteSpace(request))
                return "<empty>";

            int pipeIndex = request.IndexOf('|');
            string commandName = pipeIndex >= 0
                ? request.Substring(0, pipeIndex)
                : request;

            if (string.Equals(commandName, "ADD_ARCHIVE", StringComparison.OrdinalIgnoreCase))
            {
                var parts = request.Split(new[] { '|' }, 4);

                if (parts.Length >= 3)
                    return $"ADD_ARCHIVE|{parts[1]}|{parts[2]}|<base64 omitted>";

                return "ADD_ARCHIVE|<base64 omitted>";
            }

            if (request.Length > MaxLoggedCommandLength)
                return request.Substring(0, MaxLoggedCommandLength) + "...";

            return request;
        }

        private static string SanitizeSingleLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Unknown server error";

            return text.Replace("\r", " ").Replace("\n", " ").Trim();
        }
    }
}
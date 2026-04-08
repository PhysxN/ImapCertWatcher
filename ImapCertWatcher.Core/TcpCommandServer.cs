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
                    client.ReceiveTimeout = 60000;
                    client.SendTimeout = 60000;

                    var request = await reader.ReadLineAsync();

                    if (string.IsNullOrWhiteSpace(request))
                        return;

                    Console.WriteLine("[SERVER] CMD: " + request);

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

                    await writer.WriteLineAsync(response.Length.ToString(CultureInfo.InvariantCulture));

                    if (response.Length > 0)
                        await writer.WriteAsync(response);

                    await writer.FlushAsync();

                    Console.WriteLine("[SERVER] RESPONSE SENT");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[SERVER ERROR] " + ex.Message);
                }
            }
        }

        private static string SanitizeSingleLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "Unknown server error";

            return text.Replace("\r", " ").Replace("\n", " ").Trim();
        }
    }
}
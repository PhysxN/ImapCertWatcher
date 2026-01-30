using System;
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

        private TcpListener _listener;
        private CancellationTokenSource _cts;

        public TcpCommandServer(int port, Func<string, string> commandHandler)
        {
            _port = port;
            _commandHandler = commandHandler;
        }

        public void Start()
        {
            _cts = new CancellationTokenSource();

            _listener = new TcpListener(IPAddress.Any, _port);
            _listener.Start();

            Task.Run(AcceptLoop);
        }

        public void Stop()
        {
            try
            {
                _cts?.Cancel();
                _listener?.Stop();
            }
            catch { }
        }

        private async Task AcceptLoop()
        {
            while (!_cts.IsCancellationRequested)
            {
                try
                {
                    var client = await _listener.AcceptTcpClientAsync();
                    _ = HandleClient(client);
                }
                catch
                {
                    if (!_cts.IsCancellationRequested)
                        throw;
                }
            }
        }

        private async Task HandleClient(TcpClient client)
        {
            using (client)
            using (var stream = client.GetStream())
            using (var reader = new System.IO.StreamReader(stream, Encoding.UTF8))
            using (var writer = new System.IO.StreamWriter(stream, Encoding.UTF8) { AutoFlush = true })
            {
                try
                {
                    client.ReceiveTimeout = 10000;
                    client.SendTimeout = 10000;

                    // Читаем команду строкой
                    var request = await reader.ReadLineAsync();

                    if (string.IsNullOrWhiteSpace(request))
                        return;

                    Console.WriteLine("[SERVER] CMD RECEIVED: " + request);

                    string response;

                    try
                    {
                        response = _commandHandler(request);
                    }
                    catch (Exception ex)
                    {
                        response = "ERROR " + ex.Message;
                    }

                    // ГАРАНТИРОВАННЫЙ перевод строки
                    await writer.WriteLineAsync(response);

                    Console.WriteLine("[SERVER] RESPONSE SENT");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[SERVER] CLIENT ERROR: " + ex.Message);
                }
            }
        }
    }
}

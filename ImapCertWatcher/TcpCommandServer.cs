using System;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;

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
                var client = await _listener.AcceptTcpClientAsync();
                _ = HandleClient(client);
            }
        }

        private async Task HandleClient(TcpClient client)
        {
            using (client)
            using (var stream = client.GetStream())
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            using (var writer = new StreamWriter(stream, Encoding.UTF8) { AutoFlush = true })
            {
                try
                {
                    client.ReceiveTimeout = 60000;
                    client.SendTimeout = 60000;

                    // читаем команду
                    var request = await reader.ReadLineAsync();

                    if (string.IsNullOrWhiteSpace(request))
                        return;

                    Console.WriteLine("[SERVER] CMD: " + request);

                    string response = _commandHandler(request);

                    // ОТПРАВЛЯЕМ ВСЁ КАК ОДНУ СТРОКУ
                    await writer.WriteAsync(response);

                    Console.WriteLine("[SERVER] RESPONSE SENT");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("[SERVER ERROR] " + ex.Message);
                }
            }
        }
    }
}

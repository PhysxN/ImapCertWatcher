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
            {
                client.ReceiveTimeout = 10000;
                client.SendTimeout = 10000;

                var stream = client.GetStream();

                var buffer = new byte[4096];
                int read = await stream.ReadAsync(buffer, 0, buffer.Length);

                if (read <= 0)
                    return;

                var request = Encoding.UTF8.GetString(buffer, 0, read).Trim();

                string response;

                try
                {
                    response = _commandHandler(request);
                }
                catch (Exception ex)
                {
                    response = "ERROR: " + ex.Message;
                }

                var respBytes = Encoding.UTF8.GetBytes(response);
                await stream.WriteAsync(respBytes, 0, respBytes.Length);
            }
        }
    }
}

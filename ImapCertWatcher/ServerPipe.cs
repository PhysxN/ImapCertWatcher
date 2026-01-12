using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Server
{
    public class ServerPipe
    {
        private readonly ServerHost _server;

        public ServerPipe(ServerHost server)
        {
            _server = server;
        }

        public async void Start()
        {
            while (true)
            {
                using (var pipe = new NamedPipeServerStream(
                    "ImapCertWatcherPipe",
                    PipeDirection.InOut,
                    1,
                    PipeTransmissionMode.Message,
                    PipeOptions.Asynchronous))
                {
                    await pipe.WaitForConnectionAsync();

                    using (var reader = new StreamReader(pipe, Encoding.UTF8))
                    using (var writer = new StreamWriter(pipe, Encoding.UTF8) { AutoFlush = true })
                    {
                        var cmd = await reader.ReadLineAsync();

                        if (cmd == "FAST_CHECK")
                        {
                            await _server.RequestFastCheckAsync();
                            await writer.WriteLineAsync("OK");
                        }
                        else
                        {
                            await writer.WriteLineAsync("UNKNOWN");
                        }
                    }
                }
            }
        }
    }
}

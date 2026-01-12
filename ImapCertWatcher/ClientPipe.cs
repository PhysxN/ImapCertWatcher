using System.IO;
using System.IO.Pipes;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public static class ClientPipe
    {
        public static async Task<bool> RequestFastCheckAsync()
        {
            try
            {
                using (var pipe = new NamedPipeClientStream(
                    ".", "ImapCertWatcherPipe", PipeDirection.InOut))
                {
                    await pipe.ConnectAsync(1000);

                    using (var reader = new StreamReader(pipe))
                    using (var writer = new StreamWriter(pipe) { AutoFlush = true })
                    {
                        await writer.WriteLineAsync("FAST_CHECK");
                        var resp = await reader.ReadLineAsync();
                        return resp == "OK";
                    }
                }
            }
            catch
            {
                return false;
            }
        }
    }
}

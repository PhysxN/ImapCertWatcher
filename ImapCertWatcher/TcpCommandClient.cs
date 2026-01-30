using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public static class TcpCommandClient
    {
        public static async Task<string> SendAsync(string host, int port, string command)
        {
            using (var client = new TcpClient())
            {
                await client.ConnectAsync(host, port);

                using (var stream = client.GetStream())
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                using (var writer = new StreamWriter(stream, Encoding.UTF8) { AutoFlush = true })
                {
                    // отправляем команду
                    await writer.WriteLineAsync(command);

                    // читаем ответ строкой
                    var response = await reader.ReadLineAsync();

                    return response ?? "";
                }
            }
        }
    }
}

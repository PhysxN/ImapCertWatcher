using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public static class TcpCommandClient
    {
        public static async Task<string> SendAsync(
            string serverIp,
            int port,
            string command)
        {
            using (var client = new TcpClient())
            {
                await client.ConnectAsync(serverIp, port);

                var stream = client.GetStream();

                var data = Encoding.UTF8.GetBytes(command);
                await stream.WriteAsync(data, 0, data.Length);

                var buffer = new byte[4096];
                int read = await stream.ReadAsync(buffer, 0, buffer.Length);

                return Encoding.UTF8.GetString(buffer, 0, read);
            }
        }
    }
}

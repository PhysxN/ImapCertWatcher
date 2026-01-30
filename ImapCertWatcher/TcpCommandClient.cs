using System;
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

                    // ЧИТАЕМ ВЕСЬ ОТВЕТ ЦЕЛИКОМ
                    var response = await reader.ReadToEndAsync();

                    return response;
                }
            }
        }
        public static async Task DownloadAndOpenArchive(string host, int port, int certId)
        {
            string response = await SendAsync(host, port, "GET_ARCHIVE|" + certId);

            if (!response.StartsWith("ARCHIVE "))
                return;

            var payload = response.Substring(8);
            var parts = payload.Split(new[] { '|' }, 2);

            string fileName = parts[0];
            byte[] data = Convert.FromBase64String(parts[1]);

            string tempPath = Path.Combine(Path.GetTempPath(), fileName);

            File.WriteAllBytes(tempPath, data);

            System.Diagnostics.Process.Start(tempPath);
        }

    }


}
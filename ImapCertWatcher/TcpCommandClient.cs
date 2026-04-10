using System;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public static class TcpCommandClient
    {
        private const int ConnectTimeoutMs = 5000;
        private const int ReadTimeoutMs = 60000;
        private const int MaxResponseLength = 80_000_000; // должен покрывать base64-ответы для ZIP до 50 МБ
        private static readonly Encoding Utf8NoBom = new UTF8Encoding(false);

        public static async Task<string> SendAsync(string host, int port, string command)
        {
            if (string.IsNullOrWhiteSpace(host))
                throw new ArgumentException("Не задан адрес сервера.", nameof(host));

            if (port <= 0 || port > 65535)
                throw new ArgumentOutOfRangeException(nameof(port), "Некорректный порт сервера.");

            using (var client = new TcpClient())
            {
                var connectTask = client.ConnectAsync(host.Trim(), port);
                var connectCompleted = await Task.WhenAny(
                    connectTask,
                    Task.Delay(ConnectTimeoutMs));

                if (connectCompleted != connectTask)
                    throw new TimeoutException("Таймаут подключения к серверу.");

                await connectTask;

                client.ReceiveTimeout = ReadTimeoutMs;
                client.SendTimeout = ReadTimeoutMs;

                using (var stream = client.GetStream())
                using (var reader = new StreamReader(stream, Utf8NoBom, false, 4096, true))
                using (var writer = new StreamWriter(stream, Utf8NoBom, 4096, true)
                {
                    AutoFlush = true,
                    NewLine = "\n"
                })
                {
                    await writer.WriteLineAsync(command ?? string.Empty);

                    var lengthLineTask = reader.ReadLineAsync();
                    var lengthLineCompleted = await Task.WhenAny(
                        lengthLineTask,
                        Task.Delay(ReadTimeoutMs));

                    if (lengthLineCompleted != lengthLineTask)
                        throw new TimeoutException("Таймаут ожидания длины ответа от сервера.");

                    var lengthLine = await lengthLineTask;

                    if (string.IsNullOrWhiteSpace(lengthLine))
                        throw new IOException("Сервер не прислал длину ответа.");

                    if (!int.TryParse(lengthLine, out int length) || length < 0)
                        throw new IOException("Некорректная длина ответа сервера: " + lengthLine);

                    if (length > MaxResponseLength)
                        throw new IOException("Слишком большой ответ сервера: " + length);

                    if (length == 0)
                        return string.Empty;

                    var readTask = ReadExactAsync(reader, length);
                    var readCompleted = await Task.WhenAny(
                        readTask,
                        Task.Delay(ReadTimeoutMs));

                    if (readCompleted != readTask)
                        throw new TimeoutException("Таймаут чтения ответа от сервера.");

                    return await readTask;
                }
            }
        }

        private static async Task<string> ReadExactAsync(StreamReader reader, int length)
        {
            var buffer = new char[length];
            int offset = 0;

            while (offset < length)
            {
                int read = await reader.ReadAsync(buffer, offset, length - offset);

                if (read <= 0)
                    throw new EndOfStreamException("Соединение закрыто до получения полного ответа.");

                offset += read;
            }

            return new string(buffer);
        }
    }
}
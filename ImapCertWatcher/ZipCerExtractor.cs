using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace ImapCertWatcher.Utils
{
    public static class ZipCerExtractor
    {
        /// <summary>
        /// Извлекает все CER-файлы из ZIP в временную папку.
        /// </summary>
        public static List<string> ExtractCerFiles(string zipPath, Action<string> log)
        {
            var result = new List<string>();

            if (!File.Exists(zipPath))
            {
                log?.Invoke($"[ZIP] Файл не найден: {zipPath}");
                return result;
            }

            try
            {
                string tempDir = Path.Combine(
                    Path.GetTempPath(),
                    "ImapCertWatcher",
                    Guid.NewGuid().ToString("N")
                );

                Directory.CreateDirectory(tempDir);

                ZipFile.ExtractToDirectory(zipPath, tempDir);

                result = Directory
                    .EnumerateFiles(tempDir, "*.cer", SearchOption.AllDirectories)
                    .ToList();

                log?.Invoke($"[ZIP] Найдено CER файлов: {result.Count}");
            }
            catch (Exception ex)
            {
                log?.Invoke($"[ZIP] Ошибка распаковки {zipPath}: {ex.Message}");
            }

            return result;
        }
    }
}

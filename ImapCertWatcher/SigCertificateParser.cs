using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace ImapCertWatcher.Utils
{
    public class SigCertificateInfo
    {
        public string Fio { get; set; }
        public string SerialNumber { get; set; }
        public DateTime NotBefore { get; set; }
        public DateTime NotAfter { get; set; }
        public string SubjectRaw { get; set; }
        public string IssuerRaw { get; set; }
        public string Thumbprint { get; set; }
    }

    public static class SigCertificateParser
    {
        /// <summary>
        /// Основной метод: читает SIG / P7S файл и извлекает данные сертификата.
        /// Работает без КриптоПро (только чтение).
        /// </summary>
        public static bool TryParse(
            string sigFilePath,
            Action<string> log,
            out SigCertificateInfo certInfo)
        {
            certInfo = null;

            if (string.IsNullOrWhiteSpace(sigFilePath) || !File.Exists(sigFilePath))
            {
                log?.Invoke($"[SIG] Файл не найден: {sigFilePath}");
                return false;
            }

            try
            {
                byte[] raw = File.ReadAllBytes(sigFilePath);

                // Пытаемся загрузить как сертификат напрямую
                X509Certificate2 cert = null;

                try
                {
                    cert = new X509Certificate2(raw);
                }
                catch
                {
                    // Иногда .sig/.p7s содержит контейнер с несколькими сертификатами
                    var collection = new X509Certificate2Collection();
                    collection.Import(raw);

                    // Берём первый НЕ самоподписанный
                    cert = collection
                        .OfType<X509Certificate2>()
                        .FirstOrDefault(c => c.Subject != c.Issuer)
                        ?? collection.OfType<X509Certificate2>().FirstOrDefault();
                }

                if (cert == null)
                {
                    log?.Invoke("[SIG] Не удалось извлечь сертификат");
                    return false;
                }

                certInfo = new SigCertificateInfo
                {
                    Fio = ExtractFio(cert),
                    SerialNumber = NormalizeSerial(cert.SerialNumber),
                    NotBefore = cert.NotBefore,
                    NotAfter = cert.NotAfter,
                    SubjectRaw = cert.Subject,
                    IssuerRaw = cert.Issuer,
                    Thumbprint = cert.Thumbprint
                };

                LogCertInfo(certInfo, log);
                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"[SIG] Ошибка обработки файла '{sigFilePath}': {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Извлекает ФИО из Subject (CN=, либо SN+G).
        /// </summary>
        private static string ExtractFio(X509Certificate2 cert)
        {
            string subject = cert.Subject;

            string cn = GetSubjectValue(subject, "CN");
            if (!string.IsNullOrWhiteSpace(cn))
                return cn;

            string sn = GetSubjectValue(subject, "SN");
            string g = GetSubjectValue(subject, "G");

            if (!string.IsNullOrWhiteSpace(sn) && !string.IsNullOrWhiteSpace(g))
                return $"{sn} {g}";

            return subject; // fallback
        }

        private static string GetSubjectValue(string subject, string key)
        {
            foreach (var part in subject.Split(','))
            {
                var p = part.Trim();
                if (p.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase))
                    return p.Substring(key.Length + 1);
            }
            return "";
        }

        /// <summary>
        /// Нормализация серийного номера (убираем пробелы, приводим к верхнему регистру).
        /// </summary>
        private static string NormalizeSerial(string serial)
        {
            if (string.IsNullOrWhiteSpace(serial))
                return "";

            return serial.Replace(" ", "").ToUpperInvariant();
        }

        private static void LogCertInfo(SigCertificateInfo info, Action<string> log)
        {
            log?.Invoke("=== [SIG] Сертификат извлечён ===");
            log?.Invoke($"[SIG] ФИО            : {info.Fio}");
            log?.Invoke($"[SIG] Серийный №    : {info.SerialNumber}");
            log?.Invoke($"[SIG] Действует с  : {info.NotBefore:dd.MM.yyyy HH:mm:ss}");
            log?.Invoke($"[SIG] Действует по : {info.NotAfter:dd.MM.yyyy HH:mm:ss}");
            log?.Invoke($"[SIG] Thumbprint   : {info.Thumbprint}");
            log?.Invoke($"[SIG] Subject      : {info.SubjectRaw}");
            log?.Invoke($"[SIG] Issuer       : {info.IssuerRaw}");
            log?.Invoke("================================");
        }
    }
}

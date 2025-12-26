using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace ImapCertWatcher.Utils
{
    public class CerCertificateInfo
    {
        public string Fio { get; set; }
        public string CertNumber { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public string SubjectRaw { get; set; }
        public string Issuer { get; set; }
        public string Thumbprint { get; set; }
    }

    public static class CerCertificateParser
    {
        /// <summary>
        /// Читает файл сертификата (*.cer) и выводит информацию в лог.
        /// Работает без КриптоПро.
        /// </summary>
        public static bool TryParse(
            string cerFilePath,
            Action<string> log,
            out CerCertificateInfo certInfo)
        {
            certInfo = null;

            if (string.IsNullOrWhiteSpace(cerFilePath) || !File.Exists(cerFilePath))
            {
                log?.Invoke($"[CER] Файл не найден: {cerFilePath}");
                return false;
            }

            try
            {
                var cert = new X509Certificate2(cerFilePath);

                certInfo = new CerCertificateInfo
                {
                    Fio = ExtractFio(cert),
                    CertNumber = NormalizeSerial(cert.SerialNumber),
                    DateStart = cert.NotBefore,
                    DateEnd = cert.NotAfter,
                    SubjectRaw = cert.Subject,
                    Issuer = cert.Issuer,
                    Thumbprint = cert.Thumbprint
                };

                LogCertInfo(certInfo, log);
                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"[CER] Ошибка обработки файла '{cerFilePath}': {ex.Message}");
                return false;
            }
        }

        // ===== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ =====

        private static string ExtractFio(X509Certificate2 cert)
        {
            string subject = cert.Subject;

            // 1) CN=ФИО (самый частый случай)
            string cn = GetSubjectValue(subject, "CN");
            if (!string.IsNullOrWhiteSpace(cn))
                return cn;

            // 2) SN + G (реже)
            string sn = GetSubjectValue(subject, "SN");
            string g = GetSubjectValue(subject, "G");

            if (!string.IsNullOrWhiteSpace(sn) && !string.IsNullOrWhiteSpace(g))
                return $"{sn} {g}";

            // fallback
            return subject;
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

        private static string NormalizeSerial(string serial)
        {
            if (string.IsNullOrWhiteSpace(serial))
                return "";

            return serial.Replace(" ", "").ToUpperInvariant();
        }

        private static void LogCertInfo(CerCertificateInfo info, Action<string> log)
        {
            log?.Invoke("=== [CER] Сертификат загружен ===");
            log?.Invoke($"[CER] ФИО            : {info.Fio}");
            log?.Invoke($"[CER] Серийный №    : {info.CertNumber}");
            log?.Invoke($"[CER] Действует с  : {info.DateStart:dd.MM.yyyy HH:mm:ss}");
            log?.Invoke($"[CER] Действует по : {info.DateEnd:dd.MM.yyyy HH:mm:ss}");
            log?.Invoke($"[CER] Thumbprint   : {info.Thumbprint}");
            log?.Invoke($"[CER] Subject      : {info.SubjectRaw}");
            log?.Invoke($"[CER] Issuer       : {info.Issuer}");
            log?.Invoke("================================");
        }
    }
}

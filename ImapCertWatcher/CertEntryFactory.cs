using ImapCertWatcher.Data;
using ImapCertWatcher.Utils;

namespace ImapCertWatcher.Services
{
    public static class CertEntryFactory
    {
        public static CertEntry FromCerInfo(CerCertificateInfo info)
        {
            return new CertEntry
            {
                Fio = info.Fio,
                CertNumber = info.SerialNumber,
                DateStart = info.NotBefore,
                DateEnd = info.NotAfter
            };
        }
    }
}

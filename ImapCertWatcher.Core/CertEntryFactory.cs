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
                CertNumber = info.CertNumber,
                DateStart = info.DateStart,
                DateEnd = info.DateEnd
            };
        }
    }
}

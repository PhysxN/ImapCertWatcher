namespace ImapCertWatcher.Utils
{
    public class ServerSettings
    {
        // ===== MAIL =====

        public string MailHost { get; set; } = "";
        public int MailPort { get; set; } = 993;
        public bool MailUseSsl { get; set; } = true;

        public string MailLogin { get; set; } = "";
        public string MailPassword { get; set; } = "";

        // ===== IMAP =====

        public string ImapNewCertificatesFolder { get; set; } = "INBOX";
        public string ImapRevocationsFolder { get; set; } = "INBOX";

        // ===== DATABASE =====

        public string FirebirdDbPath { get; set; } = @"C:\DB_PHYSXN\CERTS.FDB";
        public string FbServer { get; set; } = "127.0.0.1";
        public string FbUser { get; set; } = "SYSDBA";
        public string FbPassword { get; set; } = "masterkey";
        public int FbDialect { get; set; } = 3;
        public string FbCharset { get; set; } = "UTF8";
        public bool IsDevelopment { get; set; } = false;

        // ===== SERVER LOGIC =====

        public int CheckIntervalMinutes { get; set; } = 60;
        public int NotifyDaysThreshold { get; set; } = 10;
        public bool NotifyOnlyInWorkHours { get; set; } = true;
        public bool AutoStartServer { get; set; }
        public bool MinimizeToTrayOnClose { get; set; }

        // ===== BIMOID =====

        public string BimoidAccountsKrasnoflotskaya { get; set; }
        public string BimoidAccountsPionerskaya { get; set; }

        // ✅ ВОТ СЮДА метод
        public ServerSettings Clone()
        {
            return new ServerSettings
            {
                MailHost = this.MailHost,
                MailPort = this.MailPort,
                MailUseSsl = this.MailUseSsl,
                MailLogin = this.MailLogin,
                MailPassword = this.MailPassword,

                ImapNewCertificatesFolder = this.ImapNewCertificatesFolder,
                ImapRevocationsFolder = this.ImapRevocationsFolder,

                FirebirdDbPath = this.FirebirdDbPath,
                FbServer = this.FbServer,
                FbUser = this.FbUser,
                FbPassword = this.FbPassword,
                FbDialect = this.FbDialect,
                FbCharset = this.FbCharset,

                CheckIntervalMinutes = this.CheckIntervalMinutes,
                NotifyDaysThreshold = this.NotifyDaysThreshold,
                NotifyOnlyInWorkHours = this.NotifyOnlyInWorkHours,
                AutoStartServer = this.AutoStartServer,
                MinimizeToTrayOnClose = this.MinimizeToTrayOnClose,

                BimoidAccountsKrasnoflotskaya = this.BimoidAccountsKrasnoflotskaya,
                BimoidAccountsPionerskaya = this.BimoidAccountsPionerskaya
            };
        }
    }
}
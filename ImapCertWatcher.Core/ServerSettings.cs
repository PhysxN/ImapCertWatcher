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

        public string ImapNewCertificatesFolder { get; set; } = "";
        public string ImapRevocationsFolder { get; set; } = "";
        // ===== DATABASE =====

        public string FirebirdDbPath { get; set; } = "";
        public string FbServer { get; set; } = "";
        public string FbUser { get; set; } = "";
        public string FbPassword { get; set; } = "";
        public int FbDialect { get; set; } = 3;
        public string FbCharset { get; set; } = "UTF8";
        public bool IsDevelopment { get; set; } = false;

        // ===== SERVER LOGIC =====

        public int ServerPort { get; set; } = 5050;
        public int CheckIntervalMinutes { get; set; } = 60;
        public int NotifyDaysThreshold { get; set; } = 10;
        public bool NotifyOnlyInWorkHours { get; set; } = true;
        public bool AutoStartServer { get; set; }
        public bool MinimizeToTrayOnClose { get; set; }

        // ===== BIMOID =====

        public string BimoidAccountsKrasnoflotskaya { get; set; } = "";
        public string BimoidAccountsPionerskaya { get; set; } = "";

        public string BimoidSenderExePath { get; set; } = @"BimoidBroadcastSender\BimoidBroadcastSender.exe";
        public string BimoidJobDirectory { get; set; } = "BimoidJobs";
        public string BimoidServer { get; set; } = "";
        public int BimoidPort { get; set; } = 7023;
        public string BimoidLogin { get; set; } = "";
        public string BimoidPassword { get; set; } = "";
        public int BimoidDelayBetweenMessagesMs { get; set; } = 300;


        public ServerSettings Clone()
        {
            return new ServerSettings
            {
                MailHost = MailHost,
                MailPort = MailPort,
                MailUseSsl = MailUseSsl,
                MailLogin = MailLogin,
                MailPassword = MailPassword,

                ImapNewCertificatesFolder = ImapNewCertificatesFolder,
                ImapRevocationsFolder = ImapRevocationsFolder,

                FirebirdDbPath = FirebirdDbPath,
                FbServer = FbServer,
                FbUser = FbUser,
                FbPassword = FbPassword,
                FbDialect = FbDialect,
                FbCharset = FbCharset,
                IsDevelopment = IsDevelopment,

                ServerPort = ServerPort,
                CheckIntervalMinutes = CheckIntervalMinutes,
                NotifyDaysThreshold = NotifyDaysThreshold,
                NotifyOnlyInWorkHours = NotifyOnlyInWorkHours,
                AutoStartServer = AutoStartServer,
                MinimizeToTrayOnClose = MinimizeToTrayOnClose,

                BimoidAccountsKrasnoflotskaya = BimoidAccountsKrasnoflotskaya,
                BimoidAccountsPionerskaya = BimoidAccountsPionerskaya,
                BimoidSenderExePath = BimoidSenderExePath,
                BimoidJobDirectory = BimoidJobDirectory,
                BimoidServer = BimoidServer,
                BimoidPort = BimoidPort,
                BimoidLogin = BimoidLogin,
                BimoidPassword = BimoidPassword,
                BimoidDelayBetweenMessagesMs = BimoidDelayBetweenMessagesMs,
            };
        }
    }
}
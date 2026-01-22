using System.ComponentModel;

namespace ImapCertWatcher.Utils
{
    public class AppSettings : INotifyPropertyChanged
    {
        // =========================================================
        // INotifyPropertyChanged
        // =========================================================

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // =========================================================
        // MAIL
        // =========================================================

        private string _mailHost = "";
        public string MailHost
        {
            get => _mailHost;
            set
            {
                if (_mailHost != value)
                {
                    _mailHost = value;
                    OnPropertyChanged(nameof(MailHost));
                }
            }
        }

        private int _mailPort = 993;
        public int MailPort
        {
            get => _mailPort;
            set
            {
                if (_mailPort != value)
                {
                    _mailPort = value;
                    OnPropertyChanged(nameof(MailPort));
                }
            }
        }

        private bool _mailUseSsl = true;
        public bool MailUseSsl
        {
            get => _mailUseSsl;
            set
            {
                if (_mailUseSsl != value)
                {
                    _mailUseSsl = value;
                    OnPropertyChanged(nameof(MailUseSsl));
                }
            }
        }

        private string _mailLogin = "";
        public string MailLogin
        {
            get => _mailLogin;
            set
            {
                if (_mailLogin != value)
                {
                    _mailLogin = value;
                    OnPropertyChanged(nameof(MailLogin));
                }
            }
        }

        // ⚠ Пароль не биндим напрямую (PasswordBox)
        public string MailPassword { get; set; } = "";

        // =========================================================
        // IMAP FOLDERS
        // =========================================================

        private string _imapNewCertificatesFolder = "INBOX";
        public string ImapNewCertificatesFolder
        {
            get => _imapNewCertificatesFolder;
            set
            {
                if (_imapNewCertificatesFolder != value)
                {
                    _imapNewCertificatesFolder = value;
                    OnPropertyChanged(nameof(ImapNewCertificatesFolder));
                }
            }
        }

        private string _imapRevocationsFolder = "INBOX";
        public string ImapRevocationsFolder
        {
            get => _imapRevocationsFolder;
            set
            {
                if (_imapRevocationsFolder != value)
                {
                    _imapRevocationsFolder = value;
                    OnPropertyChanged(nameof(ImapRevocationsFolder));
                }
            }
        }

        // =========================================================
        // FIREBIRD
        // =========================================================

        private string _firebirdDbPath = @"C:\DB_PHYSXN\CERTS.FDB";
        public string FirebirdDbPath
        {
            get => _firebirdDbPath;
            set
            {
                if (_firebirdDbPath != value)
                {
                    _firebirdDbPath = value;
                    OnPropertyChanged(nameof(FirebirdDbPath));
                }
            }
        }

        private string _fbServer = "127.0.0.1";
        public string FbServer
        {
            get => _fbServer;
            set
            {
                if (_fbServer != value)
                {
                    _fbServer = value;
                    OnPropertyChanged(nameof(FbServer));
                }
            }
        }

        private string _fbUser = "SYSDBA";
        public string FbUser
        {
            get => _fbUser;
            set
            {
                if (_fbUser != value)
                {
                    _fbUser = value;
                    OnPropertyChanged(nameof(FbUser));
                }
            }
        }

        // ⚠ Пароль тоже вручную
        public string FbPassword { get; set; } = "masterkey";

        private int _fbDialect = 3;
        public int FbDialect
        {
            get => _fbDialect;
            set
            {
                if (_fbDialect != value)
                {
                    _fbDialect = value;
                    OnPropertyChanged(nameof(FbDialect));
                }
            }
        }

        public string FbCharset { get; set; } = "UTF8";

        // =========================================================
        // SERVER
        // =========================================================

        private int _checkIntervalMinutes = 60;
        public int CheckIntervalMinutes
        {
            get => _checkIntervalMinutes;
            set
            {
                if (_checkIntervalMinutes != value)
                {
                    _checkIntervalMinutes = value;
                    OnPropertyChanged(nameof(CheckIntervalMinutes));
                }
            }
        }

        private int _notifyDaysThreshold = 10;
        public int NotifyDaysThreshold
        {
            get => _notifyDaysThreshold;
            set
            {
                if (_notifyDaysThreshold != value)
                {
                    _notifyDaysThreshold = value;
                    OnPropertyChanged(nameof(NotifyDaysThreshold));
                }
            }
        }

        private bool _notifyOnlyInWorkHours = true;
        public bool NotifyOnlyInWorkHours
        {
            get => _notifyOnlyInWorkHours;
            set
            {
                if (_notifyOnlyInWorkHours != value)
                {
                    _notifyOnlyInWorkHours = value;
                    OnPropertyChanged(nameof(NotifyOnlyInWorkHours));
                }
            }
        }

        // =========================================================
        // BIMOID
        // =========================================================

        private string _bimoidAccountsKrasnoflotskaya;
        public string BimoidAccountsKrasnoflotskaya
        {
            get => _bimoidAccountsKrasnoflotskaya;
            set
            {
                if (_bimoidAccountsKrasnoflotskaya != value)
                {
                    _bimoidAccountsKrasnoflotskaya = value;
                    OnPropertyChanged(nameof(BimoidAccountsKrasnoflotskaya));
                }
            }
        }

        private string _bimoidAccountsPionerskaya;
        public string BimoidAccountsPionerskaya
        {
            get => _bimoidAccountsPionerskaya;
            set
            {
                if (_bimoidAccountsPionerskaya != value)
                {
                    _bimoidAccountsPionerskaya = value;
                    OnPropertyChanged(nameof(BimoidAccountsPionerskaya));
                }
            }
        }

        // =========================================================
        // MISC / LEGACY
        // =========================================================

        public bool AutoStart { get; set; } = false;
        public bool MinimizeToTrayOnClose { get; set; } = true;
        public bool IsDevelopment { get; set; } = true;

        public string RevocationSubjectRegex { get; set; } =
            @"Сертификат\s+№\s*[A-F0-9]+.*(аннулирован|прекратил действие)";

        // =========================================================
        // LEGACY / BACKWARD COMPATIBILITY
        // =========================================================

        // ⚠ Используется в старом коде (MainWindow, Watchers)
        // Оставляем для совместимости
        private string _imapFolder = "INBOX";
        public string ImapFolder
        {
            get => _imapFolder;
            set
            {
                if (_imapFolder != value)
                {
                    _imapFolder = value;
                    OnPropertyChanged(nameof(ImapFolder));
                }
            }
        }

        // Фильтр получателя (старый функционал)
        private string _filterRecipient = "";
        public string FilterRecipient
        {
            get => _filterRecipient;
            set
            {
                if (_filterRecipient != value)
                {
                    _filterRecipient = value;
                    OnPropertyChanged(nameof(FilterRecipient));
                }
            }
        }

        // Префикс темы письма (старый функционал)
        private string _filterSubjectPrefix = "Сертификат";
        public string FilterSubjectPrefix
        {
            get => _filterSubjectPrefix;
            set
            {
                if (_filterSubjectPrefix != value)
                {
                    _filterSubjectPrefix = value;
                    OnPropertyChanged(nameof(FilterSubjectPrefix));
                }
            }
        }
    }
}

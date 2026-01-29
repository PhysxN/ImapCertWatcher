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

        public bool IsDevelopment { get; set; } = true;

        //===============
        //Клиент - сервер
        //===============

        public string ServerIp { get; set; } = "127.0.0.1";
        public int ServerPort { get; set; } = 5050;

        
    }
}

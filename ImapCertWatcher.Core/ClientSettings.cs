using System.ComponentModel;

namespace ImapCertWatcher.Utils
{
    public class ClientSettings : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        // ===== SERVER CONNECTION =====

        private string _serverIp = "127.0.0.1";
        public string ServerIp
        {
            get => _serverIp;
            set
            {
                if (_serverIp == value)
                    return;

                _serverIp = value;
                OnPropertyChanged(nameof(ServerIp));
            }
        }

        private int _serverPort = 5050;
        public int ServerPort
        {
            get => _serverPort;
            set
            {
                if (_serverPort == value)
                    return;

                _serverPort = value;
                OnPropertyChanged(nameof(ServerPort));
            }
        }

        // ===== UI =====

        private bool _darkTheme = false;
        public bool DarkTheme
        {
            get => _darkTheme;
            set
            {
                if (_darkTheme == value)
                    return;

                _darkTheme = value;
                OnPropertyChanged(nameof(DarkTheme));
            }
        }
    }
}
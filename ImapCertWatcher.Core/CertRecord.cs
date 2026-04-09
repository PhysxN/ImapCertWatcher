using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ImapCertWatcher.Models
{
    public class CertRecord : INotifyPropertyChanged
    {
        private int _id;
        private string _fio;
        private DateTime _dateStart;
        private DateTime _dateEnd;
        private string _certNumber;
        private string _fromAddress;
        private bool _isDeleted;
        private bool _hasArchiveInDb;
        private bool _hasArchive;
        private bool _isRevoked;
        public bool CanEditToken => !IsDeleted && !IsRevoked;
        private bool _isUpdatingFromServer;
        private DateTime? _revokeDate;
        private int? _tokenId;
        private TokenRecord _selectedToken;
        private string _tokenSn;
        private string _note;
        private string _building;
        private string _folderPath;
        private string _archivePath;
        private DateTime _messageDate;
        private ObservableCollection<TokenRecord> _availableTokens;

        public int Id
        {
            get => _id;
            set
            {
                if (_id == value)
                    return;

                _id = value;
                OnPropertyChanged(nameof(Id));
            }
        }

        public string Fio
        {
            get => FormatName(_fio);
            set
            {
                if (_fio == value)
                    return;

                _fio = value;
                OnPropertyChanged(nameof(Fio));
            }
        }

        public DateTime DateStart
        {
            get => _dateStart;
            set
            {
                if (_dateStart == value)
                    return;

                _dateStart = value;
                OnPropertyChanged(nameof(DateStart));
            }
        }

        public DateTime DateEnd
        {
            get => _dateEnd;
            set
            {
                if (_dateEnd == value)
                    return;

                _dateEnd = value;
                OnPropertyChanged(nameof(DateEnd));
                OnPropertyChanged(nameof(DaysLeft));
            }
        }

        public int DaysLeft
        {
            get
            {
                if (DateEnd == DateTime.MinValue)
                    return 0;

                return (DateEnd.Date - DateTime.Now.Date).Days;
            }
        }

        public ObservableCollection<TokenRecord> AvailableTokens
        {
            get => _availableTokens;
            set
            {
                if (_availableTokens == value)
                    return;

                _availableTokens = value;
                OnPropertyChanged(nameof(AvailableTokens));
            }
        }

        public string CertNumber
        {
            get => _certNumber;
            set
            {
                if (_certNumber == value)
                    return;

                _certNumber = value;
                OnPropertyChanged(nameof(CertNumber));
            }
        }

        public string FromAddress
        {
            get => _fromAddress;
            set
            {
                if (_fromAddress == value)
                    return;

                _fromAddress = value;
                OnPropertyChanged(nameof(FromAddress));
            }
        }

        public bool IsDeleted
        {
            get => _isDeleted;
            set
            {
                if (_isDeleted == value)
                    return;

                _isDeleted = value;
                OnPropertyChanged(nameof(IsDeleted));
                OnPropertyChanged(nameof(CanEditToken));
            }
        }

        public bool HasArchiveInDb
        {
            get => _hasArchiveInDb;
            set
            {
                if (_hasArchiveInDb == value)
                    return;

                _hasArchiveInDb = value;
                OnPropertyChanged(nameof(HasArchiveInDb));
            }
        }

        public bool HasArchive
        {
            get => _hasArchive;
            set
            {
                if (_hasArchive == value)
                    return;

                _hasArchive = value;
                OnPropertyChanged(nameof(HasArchive));
            }
        }

        public bool IsRevoked
        {
            get => _isRevoked;
            set
            {
                if (_isRevoked == value)
                    return;

                _isRevoked = value;
                OnPropertyChanged(nameof(IsRevoked));
                OnPropertyChanged(nameof(RevokedDisplay));
                OnPropertyChanged(nameof(CanEditToken));
            }
        }

        public bool IsUpdatingFromServer
        {
            get => _isUpdatingFromServer;
            set
            {
                if (_isUpdatingFromServer == value)
                    return;

                _isUpdatingFromServer = value;
                OnPropertyChanged(nameof(IsUpdatingFromServer));
            }
        }

        public DateTime? RevokeDate
        {
            get => _revokeDate;
            set
            {
                if (_revokeDate == value)
                    return;

                _revokeDate = value;
                OnPropertyChanged(nameof(RevokeDate));
                OnPropertyChanged(nameof(RevokedDisplay));
            }
        }

        public TokenRecord SelectedToken
        {
            get => _selectedToken;
            set
            {
                if (ReferenceEquals(_selectedToken, value))
                    return;

                _selectedToken = value;

                if (value != null)
                {
                    if (_tokenId != value.Id)
                    {
                        _tokenId = value.Id;
                        OnPropertyChanged(nameof(TokenId));
                    }

                    if (_tokenSn != value.Sn)
                    {
                        _tokenSn = value.Sn;
                        OnPropertyChanged(nameof(TokenSn));
                    }
                }
                else
                {
                    if (_tokenId != null)
                    {
                        _tokenId = null;
                        OnPropertyChanged(nameof(TokenId));
                    }

                    if (_tokenSn != null)
                    {
                        _tokenSn = null;
                        OnPropertyChanged(nameof(TokenSn));
                    }
                }

                OnPropertyChanged(nameof(SelectedToken));
            }
        }

        public int? TokenId
        {
            get => _tokenId;
            set
            {
                if (_tokenId == value)
                    return;

                _tokenId = value;
                OnPropertyChanged(nameof(TokenId));
            }
        }

        public string TokenSn
        {
            get => _tokenSn;
            set
            {
                if (_tokenSn == value)
                    return;

                _tokenSn = value;
                OnPropertyChanged(nameof(TokenSn));
            }
        }

        public string Note
        {
            get => _note;
            set
            {
                if (_note == value)
                    return;

                _note = value;
                OnPropertyChanged(nameof(Note));
            }
        }

        public string Building
        {
            get => _building;
            set
            {
                if (_building == value)
                    return;

                _building = value;
                OnPropertyChanged(nameof(Building));
            }
        }

        public string RevokedDisplay
        {
            get
            {
                if (!IsRevoked)
                    return "";

                return RevokeDate.HasValue
                    ? $"Да ({RevokeDate.Value:dd.MM.yyyy})"
                    : "Да";
            }
        }

        public string FolderPath
        {
            get => _folderPath;
            set
            {
                if (_folderPath == value)
                    return;

                _folderPath = value;
                OnPropertyChanged(nameof(FolderPath));
            }
        }

        public string ArchivePath
        {
            get => _archivePath;
            set
            {
                if (_archivePath == value)
                    return;

                _archivePath = value;
                OnPropertyChanged(nameof(ArchivePath));
            }
        }

        public DateTime MessageDate
        {
            get => _messageDate;
            set
            {
                if (_messageDate == value)
                    return;

                _messageDate = value;
                OnPropertyChanged(nameof(MessageDate));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void RefreshDaysLeft()
        {
            OnPropertyChanged(nameof(DaysLeft));
        }

        private string FormatName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var text = value.Trim().ToLower();
            var words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < words.Length; i++)
            {
                if (words[i].Length > 1)
                    words[i] = char.ToUpper(words[i][0]) + words[i].Substring(1);
                else
                    words[i] = words[i].ToUpper();
            }

            return string.Join(" ", words);
        }
    }
}
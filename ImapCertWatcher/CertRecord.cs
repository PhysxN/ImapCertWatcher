using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace ImapCertWatcher.Models
{
    public class CertRecord : INotifyPropertyChanged
    {
        private string _building;
        private string _note;
        private DateTime _dateEnd;

        public int Id { get; set; }
        public string Fio { get; set; }
        public DateTime DateStart { get; set; }

        public DateTime DateEnd
        {
            get => _dateEnd;
            set
            {
                if (_dateEnd != value)
                {
                    _dateEnd = value;
                    OnPropertyChanged(nameof(DateEnd));
                    OnPropertyChanged(nameof(DaysLeft)); // Обновляем DaysLeft при изменении DateEnd
                }
            }
        }

        // Убираем сохраненное значение DaysLeft, вычисляем динамически
        public int DaysLeft
        {
            get
            {
                if (DateEnd == DateTime.MinValue)
                    return 0;

                return (DateEnd.Date - DateTime.Now.Date).Days;
            }
        }

        private ObservableCollection<TokenRecord> _availableTokens;

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

        public string CertNumber { get; set; }
        public string FromAddress { get; set; }
        public bool IsDeleted { get; set; }
        public bool HasArchiveInDb { get; set; }
        public bool HasArchive { get; set; }
        public bool IsRevoked { get; set; }
        public bool IsUpdatingFromServer { get; set; }
        public DateTime? RevokeDate { get; set; }
        private int? _tokenId;
        public int? TokenId
        {
            get => _tokenId;
            set
            {
                // ⛔ ГЛАВНЫЙ СТОП-ПЕТЛЯ
                if (_tokenId == value)
                    return;

                _tokenId = value;
                OnPropertyChanged(nameof(TokenId));
            }
        }
        public string TokenSn { get; set; }
        public string Note
        {
            get => _note;
            set
            {
                if (_note != value)
                {
                    _note = value;
                    OnPropertyChanged(nameof(Note));
                }
            }
        }

        public string Building
        {
            get => _building;
            set
            {
                if (_building != value)
                {
                    _building = value;
                    OnPropertyChanged(nameof(Building));
                }
            }
        }

        public string RevokedDisplay
        {
            get
            {
                if (!IsRevoked) return "";
                return RevokeDate.HasValue
                    ? $"Да ({RevokeDate.Value:dd.MM.yyyy})"
                    : "Да";
            }
        }

        public string FolderPath { get; set; }
        public string ArchivePath { get; set; }
        public DateTime MessageDate { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // ★ ДОБАВЛЯЕМ ПУБЛИЧНЫЙ МЕТОД ДЛЯ ОБНОВЛЕНИЯ ★
        public void RefreshDaysLeft()
        {
            OnPropertyChanged(nameof(DaysLeft));
        }


    }
}
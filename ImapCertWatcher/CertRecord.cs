using System;
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
                if (DateEnd == DateTime.MinValue) return 0;
                var days = (int)(DateEnd - DateTime.Now).TotalDays;
                return Math.Max(0, days); // Не показываем отрицательные значения
            }
        }

        public string CertNumber { get; set; }
        public string FromAddress { get; set; }
        public bool IsDeleted { get; set; }
        public bool HasArchiveInDb { get; set; }
        public bool HasArchive { get; set; }

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

        public string FolderPath { get; set; }
        public string ArchivePath { get; set; }
        public DateTime MessageDate { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
using System;

namespace ImapCertWatcher.Models
{
    public class CertRecord
    {
        public int Id { get; set; }
        public string Fio { get; set; }
        public System.DateTime DateStart { get; set; }
        public System.DateTime DateEnd { get; set; }
        public int DaysLeft { get; set; }
        public string CertNumber { get; set; }
        public string FromAddress { get; set; }
        public bool IsDeleted { get; set; }
        public string Note { get; set; }
        public string Building { get; set; }
        public string FolderPath { get; set; }
        public string ArchivePath { get; set; } // Новое поле - путь к архиву
        public DateTime MessageDate { get; set; } // Новое поле - дата письма для определения самого нового
    }
}
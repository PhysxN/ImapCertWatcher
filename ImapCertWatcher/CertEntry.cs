using System;

namespace ImapCertWatcher.Models
{
    public class CertEntry
    {
        public int Id { get; set; }
        public string MailUid { get; set; }
        public string Fio { get; set; }
        public DateTime DateStart { get; set; }
        public DateTime DateEnd { get; set; }
        public int DaysLeft { get; set; }
        public string Subject { get; set; }
        public DateTime Received { get; set; }
        public string CertNumber { get; set; }
        public string FromAddress { get; set; }
        public string FolderPath { get; set; }
        public string ArchivePath { get; set; } // Новое поле - путь к архиву
        public DateTime MessageDate { get; set; } // Новое поле - дата письма
    }
}
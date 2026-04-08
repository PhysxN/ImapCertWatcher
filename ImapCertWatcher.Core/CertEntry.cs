using System;

namespace ImapCertWatcher.Data
{
    public class CertEntry
    {
        // Идентификатор (если используется)
        public int Id { get; set; }

        // Основные поля, которые используются в Watcher'ах и DbHelper
        public string Fio { get; set; }
        public string CertNumber { get; set; }
        public DateTime DateStart { get; set; } = DateTime.MinValue;
        public DateTime DateEnd { get; set; } = DateTime.MinValue;
        public int DaysLeft { get; set; }
        public string Subject { get; set; }
        public DateTime MessageDate { get; set; }
        public DateTime Received { get; set; }
        public string FromAddress { get; set; }
        public string FolderPath { get; set; }
        public string MailUid { get; set; }

        // Новые поля, которые требуются — добавлены:
        public string Note { get; set; }          // поле Примечание
        public string Building { get; set; }      // поле Здание

        // Поля для архива/пути к архиву (если используются)
        public string ArchivePath { get; set; }
        public bool HasArchiveInDb { get; set; }

        // Дополнительные служебные поля
        public bool IsDeleted { get; set; }
    }
}

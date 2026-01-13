using System;

namespace ImapCertWatcher.Utils
{
    public class AppSettings
    {
        // Почта
        public string MailHost { get; set; } = "";
        public int MailPort { get; set; } = 993;
        public bool MailUseSsl { get; set; } = true;
        public string MailLogin { get; set; } = "";
        public string MailPassword { get; set; } = "";
        public string ImapFolder { get; set; } = "INBOX";
        // Папка с письмами, содержащими ZIP/CER с новыми сертификатами
        public string ImapNewCertificatesFolder { get; set; }
            = "INBOX";
        public int NotifyDaysThreshold { get; set; } = 10;          // сколько дней до окончания, чтобы слать
        public string BimoidAccountsKrasnoflotskaya { get; set; }   // аккаунты для Краснофлотской (по одному в строке)
        public string BimoidAccountsPionerskaya { get; set; }       // аккаунты для Пионерской (по одному в строке)

        // Фильтры для поиска писем
        public string FilterRecipient { get; set; } = "";
        // Префикс темы, по которому определяется "новое" письмо (по умолчанию "Сертификат №" или "Сертификат")
        public string FilterSubjectPrefix { get; set; } = "Сертификат";
        //Разработка
        public bool IsDevelopment { get; set; } = true;
        // БД Firebird
        public string FirebirdDbPath { get; set; }
    = @"C:\DB_PHYSXN\CERTS.FDB";
        public string FbServer { get; set; }
    = "127.0.0.1";
        public string FbUser { get; set; } = "SYSDBA";
        public string FbPassword { get; set; } = "masterkey";
        public int FbDialect { get; set; } = 3;        
        // Автозагрузка и трей
        public bool AutoStart { get; set; } = false;
        public bool MinimizeToTrayOnClose { get; set; } = true;

        // Отправлять уведомления / запускать авто-проверку только в рабочие часы
        public bool NotifyOnlyInWorkHours { get; set; } = true;

        // Кодировка (если DbHelper ожидает поле FbCharset)
        public string FbCharset { get; set; } = "UTF8";

        // Интервал проверки почты (в часах)
        public int CheckIntervalMinutes { get; set; } = 60; // по умолчанию 60 минут

        // Любые дополнительные настройки, которые захотите
        public bool SomeFeatureToggle { get; set; } = false;

        public string ImapServer
        {
            get => MailHost;
            set => MailHost = value;
        }

        public int ImapPort
        {
            get => MailPort;
            set => MailPort = value;
        }

        public string MailUser
        {
            get => MailLogin;
            set => MailLogin = value;
        }

        // Папка с письмами об аннулировании (если хотите отличную от общего ImapFolder)
        public string ImapRevocationsFolder { get; set; } = "INBOX";

        // Регекс для строгой проверки темы аннулирований (можно поменять)
        public string RevocationSubjectRegex { get; set; } =
            @"Сертификат\s+№\s*[A-F0-9]+.*(аннулирован|прекратил действие)";
    }
}

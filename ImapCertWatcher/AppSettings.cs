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

        public int NotifyDaysThreshold { get; set; } = 10;          // сколько дней до окончания, чтобы слать
        public string BimoidAccountsKrasnoflotskaya { get; set; }   // аккаунты для Краснофлотской (по одному в строке)
        public string BimoidAccountsPionerskaya { get; set; }       // аккаунты для Пионерской (по одному в строке)

        // Фильтры для поиска писем
        public string FilterRecipient { get; set; } = "";
        // Префикс темы, по которому определяется "новое" письмо (по умолчанию "Сертификат №" или "Сертификат")
        public string FilterSubjectPrefix { get; set; } = "Сертификат";

        // БД Firebird
        public string FirebirdDbPath { get; set; } = "";
        public string FbServer { get; set; } = "127.0.0.1";
        public string FbUser { get; set; } = "SYSDBA";
        public string FbPassword { get; set; } = "masterkey";
        public int FbDialect { get; set; } = 3;
        // Кодировка (если DbHelper ожидает поле FbCharset)
        public string FbCharset { get; set; } = "UTF8";

        // Интервал проверки почты (в часах)
        public int CheckIntervalHours { get; set; } = 1;

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

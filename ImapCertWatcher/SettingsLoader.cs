using System;
using System.IO;
using System.Text;

namespace ImapCertWatcher.Utils
{
    public static class SettingsLoader
    {
        public static AppSettings Load(string path)
        {
            var s = new AppSettings();
            if (!File.Exists(path))
                throw new FileNotFoundException("settings.txt not found", path);

            foreach (var raw in File.ReadAllLines(path, Encoding.UTF8))
            {
                var line = raw.Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#")) continue;
                var idx = line.IndexOf('=');
                if (idx <= 0) continue;
                var key = line.Substring(0, idx).Trim();
                var val = line.Substring(idx + 1).Trim();

                if (string.IsNullOrEmpty(val)) continue;

                switch (key)
                {
                    case "MailHost": s.MailHost = val; break;
                    case "MailPort": if (int.TryParse(val, out int mp)) s.MailPort = mp; break;
                    case "MailUseSsl": if (bool.TryParse(val, out bool mus)) s.MailUseSsl = mus; break;
                    case "MailLogin": s.MailLogin = val; break;
                    case "MailPassword": s.MailPassword = val; break;
                    case "ImapFolder": s.ImapFolder = val; break;
                    case "ImapNewCertificatesFolder":
                        s.ImapNewCertificatesFolder = val;
                        break;

                    case "ImapRevocationsFolder":
                        s.ImapRevocationsFolder = val;
                        break;
                    case "FilterRecipient": s.FilterRecipient = val; break;
                    case "FilterSubjectPrefix": s.FilterSubjectPrefix = val; break;

                    case "FirebirdDbPath": s.FirebirdDbPath = val; break;
                    case "FbServer": s.FbServer = val; break;
                    case "FbUser": s.FbUser = val; break;
                    case "FbPassword": s.FbPassword = val; break;
                    case "FbDialect": if (int.TryParse(val, out int d)) s.FbDialect = d; break;
                    case "FbCharset": s.FbCharset = val; break;

                    // Интервал проверки: ожидаем минуты.
                    // Поддерживаем старый ключ "CheckIntervalHours" (legacy) — в этом случае умножаем на 60.
                    case "CheckIntervalMinutes":
                        if (int.TryParse(val, out int cm)) s.CheckIntervalMinutes = cm;
                        break;
                    case "CheckIntervalHours":
                        if (int.TryParse(val, out int ch)) s.CheckIntervalMinutes = ch * 60;
                        break;

                    // === НОВЫЕ НАСТРОЙКИ ДЛЯ ОБИМПА ===
                    case "NotifyDaysThreshold":
                        if (int.TryParse(val, out int nd))
                            s.NotifyDaysThreshold = nd;
                        break;

                    case "BimoidAccountsKrasnoflotskaya":
                        // \n в файле превращаем в реальные переводы строки
                        s.BimoidAccountsKrasnoflotskaya = val.Replace("\\n", Environment.NewLine);
                        break;

                    case "BimoidAccountsPionerskaya":
                        s.BimoidAccountsPionerskaya = val.Replace("\\n", Environment.NewLine);
                        break;

                    case "AutoStart":
                        if (bool.TryParse(val, out bool auto))
                            s.AutoStart = auto;
                        break;

                    case "MinimizeToTrayOnClose":
                        if (bool.TryParse(val, out bool mt))
                            s.MinimizeToTrayOnClose = mt;
                        break;

                    case "NotifyOnlyInWorkHours":
                        if (bool.TryParse(val, out bool nw))
                            s.NotifyOnlyInWorkHours = nw;
                        break;
                    case "ServerIp":
                        s.ServerIp = val;
                        break;

                    case "ServerPort":
                        if (int.TryParse(val, out int sp))
                            s.ServerPort = sp;
                        break;
                }
            }

            // если NotifyDaysThreshold вдруг не прочитался – оставим дефолт 10
            if (s.NotifyDaysThreshold <= 0)
                s.NotifyDaysThreshold = 10;

            return s;
        }
    }
}

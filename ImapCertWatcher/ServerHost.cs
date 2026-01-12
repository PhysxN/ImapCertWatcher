using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.Threading.Tasks;

namespace ImapCertWatcher.Server
{
    public class ServerHost : IDisposable
    {
        private readonly SafeAsyncTimer _timer;
        private readonly ImapNewCertificatesWatcher _newWatcher;
        private readonly ImapRevocationsWatcher _revokeWatcher;
        private readonly NotificationManager _notifications;
        private readonly DbHelper _db;

        public ServerHost(AppSettings settings, DbHelper db, Action<string> log)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));

            _newWatcher = new ImapNewCertificatesWatcher(settings, _db, log);
            _revokeWatcher = new ImapRevocationsWatcher(settings, _db, log);
            _notifications = new NotificationManager(settings, log);

            _timer = new SafeAsyncTimer(
                async ct =>
                {
                    log("ServerHost: проверка почты (FAST)");

                    var checkStartedAt = DateTime.Now;

                    await _newWatcher.CheckNewCertificatesFastAsync();
                    await _revokeWatcher.CheckRevocationsFastAsync();

                    var newCerts = _db.GetCertificatesAddedAfter(checkStartedAt);
                    if (newCerts.Count > 0)
                    {
                        _notifications.NotifyNewUsers(newCerts);
                    }
                },
                ex => log("ServerHost error: " + ex.Message)
            );
        }

        public void Start()
        {
            // первая проверка через 10 секунд, далее каждые 10 минут
            _timer.Start(
                initialDelay: TimeSpan.FromSeconds(10),
                period: TimeSpan.FromMinutes(10)
            );
        }

        public async Task RequestFastCheckAsync()
        {
            var checkStartedAt = DateTime.Now;

            await _newWatcher.CheckNewCertificatesFastAsync();
            await _revokeWatcher.CheckRevocationsFastAsync();

            var newCerts = _db.GetCertificatesAddedAfter(checkStartedAt);
            if (newCerts.Count > 0)
            {
                _notifications.NotifyNewUsers(newCerts);
            }
        }

        public void Dispose()
        {
            _timer.Dispose();
        }
    }
}

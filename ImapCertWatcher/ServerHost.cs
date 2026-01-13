using ImapCertWatcher.Data;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.IO;
using System.IO.Pipes;
using System.Text;
using System.Threading;
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
        private readonly Action<string> _log;
        private CancellationTokenSource _pipeCts;
        private Task _pipeTask;

        private const string PipeName = "ImapCertWatcherPipe";
        public ServerHost(AppSettings settings, DbHelper db, Action<string> log)
        {
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });   // ← 🔥 ВОТ ЭТО ОБЯЗАТЕЛЬНО

            _newWatcher = new ImapNewCertificatesWatcher(settings, _db, _log);
            _revokeWatcher = new ImapRevocationsWatcher(settings, _db, _log);
            _notifications = new NotificationManager(settings, _log);


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
            _timer.Start(
                initialDelay: TimeSpan.FromSeconds(10),
                period: TimeSpan.FromMinutes(10)
            );

            StartPipeServer(); // ← ВОТ ЭТО
        }

        private void StartPipeServer()
        {
            _pipeCts = new CancellationTokenSource();

            _pipeTask = Task.Run(async () =>
            {
                while (!_pipeCts.IsCancellationRequested)
                {
                    try
                    {
                        using (var server = new NamedPipeServerStream(
                            PipeName,
                            PipeDirection.In,
                            1,
                            PipeTransmissionMode.Message,
                            PipeOptions.Asynchronous))
                        {
                            _log("Pipe: ожидание команды клиента...");
                            await server.WaitForConnectionAsync(_pipeCts.Token);

                            using (var reader = new StreamReader(server, Encoding.UTF8))
                            {
                                var command = await reader.ReadLineAsync();
                                _log($"Pipe: получена команда '{command}'");

                                if (command == "FAST_CHECK")
                                {
                                    await RequestFastCheckAsync();
                                }
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        // корректное завершение
                        break;
                    }
                    catch (Exception ex)
                    {
                        _log("Pipe error: " + ex.Message);
                        await Task.Delay(1000);
                    }
                }
            });
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
            _pipeCts?.Cancel();
            _timer.Dispose();
        }
    }
}

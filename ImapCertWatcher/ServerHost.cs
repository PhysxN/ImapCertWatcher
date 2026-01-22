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
        private const string PipeName = "ImapCertWatcherPipe";
        private readonly AppSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        private readonly SafeAsyncTimer _timer;
        private readonly ImapNewCertificatesWatcher _newWatcher;
        private readonly ImapRevocationsWatcher _revokeWatcher;
        private readonly NotificationManager _notifications;

        private CancellationTokenSource _pipeCts;
        private Task _pipeTask;

        public AppSettings Settings => _settings;

        public ServerHost(AppSettings settings, DbHelper db, Action<string> log)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _db = db ?? throw new ArgumentNullException(nameof(db));
            _log = log ?? (_ => { });

            _newWatcher = new ImapNewCertificatesWatcher(_settings, _db, _log);
            _revokeWatcher = new ImapRevocationsWatcher(_settings, _db, _log);
            _notifications = new NotificationManager(_settings, _log);

            _timer = new SafeAsyncTimer(
                async ct =>
                {
                    _log("ServerHost: проверка почты (FAST)");

                    var checkStartedAt = DateTime.Now;

                    await _newWatcher.CheckNewCertificatesFastAsync();
                    await _revokeWatcher.CheckRevocationsFastAsync();

                    var newCerts = _db.GetCertificatesAddedAfter(checkStartedAt);
                    if (newCerts.Count > 0)
                        _notifications.NotifyNewUsers(newCerts);
                },
                ex => _log("ServerHost error: " + ex.Message)
            );
        }


        public void Start()
        {
            if (_settings.CheckIntervalMinutes > 0)
            {
                var interval = TimeSpan.FromMinutes(_settings.CheckIntervalMinutes);

                _timer.Start(
                    initialDelay: interval,
                    period: interval
                );

                _log($"ServerHost: таймер запущен, интервал {_settings.CheckIntervalMinutes} мин.");
            }
            else
            {
                _log("ServerHost: авто-проверка отключена (CheckIntervalMinutes <= 0)");
            }

            StartPipeServer();
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

        public Task RequestFullCheckAsync()
        {
            return Task.Run(() =>
            {
                _log("ПОЛНАЯ проверка почты");

                // ✅ новые сертификаты — ВСЕ письма
                _newWatcher.ProcessNewCertificates(checkAllMessages: true);

                // ✅ аннулирования — ВСЕ письма
                _revokeWatcher.ProcessRevocations(checkAllMessages: true);

                _log("ПОЛНАЯ проверка завершена");
            });
        }

        public void Dispose()
        {
            _pipeCts?.Cancel();
            _pipeTask?.Wait(1000);
            _timer.Dispose();
        }
    }
}

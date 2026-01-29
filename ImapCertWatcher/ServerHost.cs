using ImapCertWatcher.Data;
using ImapCertWatcher.Server;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;


namespace ImapCertWatcher.Server
{
    public class ServerHost : IDisposable
    {
        private const string PipeName = "ImapCertWatcherPipe";
        private readonly ServerSettings _settings;
        private readonly DbHelper _db;
        private readonly Action<string> _log;

        private readonly SafeAsyncTimer _timer;
        private readonly ImapNewCertificatesWatcher _newWatcher;
        private readonly ImapRevocationsWatcher _revokeWatcher;
        private readonly NotificationManager _notifications;

        private CancellationTokenSource _pipeCts;
        private Task _pipeTask;
        private TcpCommandServer _tcpServer;

        public ServerSettings Settings => _settings;



        private readonly SemaphoreSlim _checkLock = new SemaphoreSlim(1, 1);

        private DateTime _lastCheckTime = DateTime.MinValue;

        private readonly TimeSpan _minCheckInterval = TimeSpan.FromMinutes(10);
        private volatile int _progressPercent = 0;
        private volatile string _currentStage = "IDLE";
        private readonly object _logLock = new object();

        private DateTime _nextTimerRun = DateTime.MinValue;

        public ServerHost(ServerSettings settings, DbHelper db, Action<string> log)
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
                            _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);
                        },
                        ex => _log("ServerHost timer error: " + ex.Message)
                        );
        }


        public void Start()
        {
            if (_settings.CheckIntervalMinutes > 0)
            {
                var interval = TimeSpan.FromMinutes(_settings.CheckIntervalMinutes);

                _nextTimerRun = DateTime.Now.Add(interval);

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
            _tcpServer = new TcpCommandServer(5050, HandleRemoteCommand);
            _tcpServer.Start();
            _log("TCP Server started. Listening on port 5050");
        }

        private string HandleRemoteCommand(string cmd)
        {
            cmd = cmd.Trim().ToUpperInvariant();

            switch (cmd)
            {
                case "FAST_CHECK":
                    _ = Task.Run(() => RequestFastCheckAsync(CancellationToken.None));
                    return "OK FAST STARTED";

                case "FULL_CHECK":
                    _ = Task.Run(() => RequestFullCheckAsync());
                    return "OK FULL STARTED";

                case "STATUS":
                    return _checkLock.CurrentCount == 0
                        ? "STATUS BUSY"
                        : "STATUS IDLE";

                case "GET_PROGRESS":
                    return $"PROGRESS {_progressPercent}|{_currentStage}";

                case "GET_TIMER":
                    var remain = (_nextTimerRun - DateTime.Now);
                    if (remain.TotalSeconds < 0)
                        return "TIMER READY";

                    return $"TIMER {(int)remain.TotalMinutes + 1}";

                case "GET_LOG":
                    return "LOG " + GetLastServerLogLines();

                default:
                    return "UNKNOWN COMMAND";
            }
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
                                    await RequestFastCheckAsync(CancellationToken.None);
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

        public async Task<string> RequestFastCheckAsync(CancellationToken token)
        {
            // защита частоты
            var sinceLast = DateTime.Now - _lastCheckTime;

            if (sinceLast < _minCheckInterval)
            {
                var wait = _minCheckInterval - sinceLast;
                return $"BUSY WAIT {(int)wait.TotalMinutes + 1} MIN";
            }

            // защита параллельности
            if (!await _checkLock.WaitAsync(0))
            {
                return "BUSY RUNNING";
            }

            try
            {
                _lastCheckTime = DateTime.Now;

                _log("FAST CHECK started");
                _progressPercent = 5;
                _currentStage = "Connecting";

                var startedAt = DateTime.Now;

                _progressPercent = 30;
                _currentStage = "Checking mail";
                await Task.WhenAll(
                    _newWatcher.CheckNewCertificatesFastAsync(token),
                    _revokeWatcher.CheckRevocationsFastAsync(token)
                );

                var newCerts = _db.GetCertificatesAddedAfter(startedAt);

                _progressPercent = 80;
                _currentStage = "Saving data";
                if (newCerts.Count > 0)
                    _notifications.NotifyNewUsers(newCerts);

                _log("FAST CHECK finished");
                _progressPercent = 100;
                _currentStage = "IDLE";

                return "OK FAST DONE";
            }
            finally
            {
                _checkLock.Release();
            }
        }

        public Task RequestFullCheckAsync()
        {
            return Task.Run(() =>
            {
                _log("ПОЛНАЯ проверка почты");

                // ✅ новые сертификаты — ВСЕ письма
                _newWatcher.ProcessNewCertificates(true, CancellationToken.None);

                // ✅ аннулирования — ВСЕ письма
                _revokeWatcher.ProcessRevocations(true, CancellationToken.None);

                _log("ПОЛНАЯ проверка завершена");
            });
        }

        private string GetLastServerLogLines(int count = 20)
        {
            try
            {
                var logFile = LogSession.SessionLogFile;

                if (!File.Exists(logFile))
                    return "";

                var lines = File.ReadAllLines(logFile);

                return string.Join("\n",
                    lines.Skip(Math.Max(0, lines.Length - count)));
            }
            catch
            {
                return "";
            }
        }
        public void Dispose()
        {
            _pipeCts?.Cancel();
            _pipeTask?.Wait(1000);
            _timer.Dispose();
        }
    }
}

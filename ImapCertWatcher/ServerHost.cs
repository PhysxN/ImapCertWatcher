using ImapCertWatcher.Data;
using ImapCertWatcher.Server;
using ImapCertWatcher.Services;
using ImapCertWatcher.Utils;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;


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

        
        
        private TcpCommandServer _tcpServer;

        public ServerSettings Settings => _settings;



        private readonly SemaphoreSlim _checkLock = new SemaphoreSlim(1, 1);

        private DateTime _lastCheckTime = DateTime.MinValue;

        private readonly TimeSpan _minCheckInterval = TimeSpan.FromMinutes(10);
        private volatile int _progressPercent = 0;
        private volatile string _currentStage = "IDLE";
        private readonly object _stateLock = new object();
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

            _tcpServer = new TcpCommandServer(5050, HandleRemoteCommand);
            _tcpServer.Start();

            _log("TCP Server started. Listening on port 5050");
        }

        private string HandleRemoteCommand(string cmd)
        {
            var originalCmd = cmd.Trim();
            var upperCmd = originalCmd.ToUpperInvariant();

            var command = upperCmd.Split('|')[0];
            if (!command.StartsWith("GET_") &&
    command != "STATUS")
            {
                _log("CMD IN: " + originalCmd);
            }

            switch (command)
            {
                case "FAST_CHECK":

                    if (_checkLock.CurrentCount == 0)
                        return "BUSY";

                    _ = Task.Run(async () =>
                    {
                        await RequestFastCheckAsync(CancellationToken.None);
                    });

                    return "OK";

                case "FULL_CHECK":
                    _ = Task.Run(() => RequestFullCheckAsync());
                    return "OK FULL STARTED";

                case "STATUS":

                    if (_checkLock.CurrentCount == 0)
                        return "STATUS BUSY";

                    return "STATUS IDLE";

                case "GET_PROGRESS":
                    lock (_stateLock)
                        return $"PROGRESS {_progressPercent}|{_currentStage}";

                case "GET_TIMER":
                    var remain = (_nextTimerRun - DateTime.Now);
                    if (remain.TotalSeconds < 0)
                        return "TIMER READY";

                    return $"TIMER {(int)remain.TotalMinutes + 1}";

                case "GET_LOG":
                    return "LOG " + GetLastServerLogLines();

                case "GET_CERTS":
                    try
                    {
                        var list = _db.GetAllCertificates();

                        var json = JsonConvert.SerializeObject(list);

                        _log("GET_CERTS -> send " + list.Count + " records");

                        // ВАЖНО: добавляем перевод строки!
                        return "CERTS " + json + "\n";
                    }
                    catch (Exception ex)
                    {
                        _log("GET_CERTS error: " + ex.Message);
                        return "CERTS []\n";
                    }

                case "MARK_DELETED":

                    var dparts = cmd.Split('|');

                    int did = int.Parse(dparts[1]);
                    bool deleted = bool.Parse(dparts[2]);

                    _db.MarkAsDeleted(did, deleted);
                    _log($"MARK_DELETED OK: ID={did}, Deleted={deleted}");

                    return "OK DELETED UPDATED\n";

                case "UPDATE_NOTE":
                    {
                        try
                        {
                            var parts = cmd.Split(new[] { '|' }, 3);

                            int nid = int.Parse(parts[1]);
                            string note = parts[2];

                            _log($"UPDATE_NOTE received: ID={nid}");

                            _db.UpdateNote(nid, note);

                            _log($"UPDATE_NOTE OK: ID={nid}");

                            return "OK NOTE UPDATED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("UPDATE_NOTE ERROR: " + ex.Message);
                            return "ERROR NOTE UPDATE\n";
                        }
                    }

                case "ADD_ARCHIVE":
                    {
                        try
                        {
                            var aparts = cmd.Split(new[] { '|' }, 4);

                            int certId = int.Parse(aparts[1]);
                            string certNumber = aparts[2];
                            byte[] data = Convert.FromBase64String(aparts[3]);

                            _log($"ADD_ARCHIVE received: ID={certId}, Cert={certNumber}, Size={data.Length} bytes");

                            string safeName = certNumber.Replace(" ", "_").Replace("/", "_");
                            string fileName = safeName + ".zip";

                            _db.SaveArchiveToDb(certId, fileName, data);

                            _log($"ADD_ARCHIVE OK: ID={certId}");

                            return "OK ARCHIVE SAVED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("ADD_ARCHIVE ERROR: " + ex.Message);
                            return "ERROR ARCHIVE SAVE\n";
                        }
                    }

                case "GET_ARCHIVE":
                    {

                        int gid = int.Parse(cmd.Split('|')[1]);

                        var result = _db.GetArchiveFromDb(gid);

                        if (result.data == null)
                            return "ARCHIVE EMPTY\n";
                        _log($"GET_ARCHIVE request: ID={gid}");
                        string fileName = result.fileName;
                        var base64 = Convert.ToBase64String(result.data);
                        _log($"GET_ARCHIVE OK: ID={gid}, Size={result.data.Length}");
                        return "ARCHIVE|" + fileName + ".zip|" + base64;
                    }

                case "SET_BUILDING":
                    {
                        try
                        {
                            var bparts = originalCmd.Split('|');

                            int bid = int.Parse(bparts[1]);
                            string building = bparts[2];

                            _log($"SET_BUILDING received: ID={bid}, Building='{building}'");

                            _db.UpdateBuilding(bid, building);

                            _log($"SET_BUILDING OK: ID={bid}");

                            return "OK BUILDING UPDATED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("SET_BUILDING ERROR: " + ex.Message);
                            return "ERROR BUILDING UPDATE\n";
                        }
                    }


                default:
                    _log("UNKNOWN CMD: " + originalCmd);
                    return "UNKNOWN COMMAND\n";
            }
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
                lock (_stateLock)
                {
                    _progressPercent = 5;
                    _currentStage = "Connecting";
                }

                var startedAt = DateTime.Now;
                lock (_stateLock)
                {
                    _progressPercent = 30;
                _currentStage = "Checking mail";
                }
                await Task.WhenAll(
                    _newWatcher.CheckNewCertificatesFastAsync(token),
                    _revokeWatcher.CheckRevocationsFastAsync(token)
                );

                var newCerts = _db.GetCertificatesAddedAfter(startedAt);
                lock (_stateLock)
                {
                    _progressPercent = 80;
                _currentStage = "Saving data";
                }
                if (newCerts.Count > 0)
                    _notifications.NotifyNewUsers(newCerts);

                _log("FAST CHECK finished");
                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

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
            _timer.Dispose();
        }
    }
}

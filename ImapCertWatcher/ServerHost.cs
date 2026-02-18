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

        private TimeSpan MinCheckInterval => TimeSpan.FromMinutes(_settings.CheckIntervalMinutes);
        private volatile int _progressPercent = 0;
        private volatile string _currentStage = "IDLE";
        private readonly object _stateLock = new object();
        

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
                        try
                        {
                            _log("AUTO CHECK triggered");

                            var result = await RequestFastCheckAsync(ct);

                            _log("AUTO CHECK result: " + result);
                        }
                        catch (Exception ex)
                        {
                            _log("AUTO CHECK ERROR: " + ex.Message);
                        }
                        finally
                        {
                            _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);
                        }
                    },
                    ex => _log("ServerHost timer fatal error: " + ex.Message)
);
        }


        public void Start()
        {
            if (_tcpServer != null)
                return;

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

            _tcpServer = new TcpCommandServer(5050, HandleRemoteCommand);
            _tcpServer.Start();
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
                    _ = RequestFastCheckAsync(CancellationToken.None);
                    return "OK";

                case "FULL_CHECK":
                    _ = RequestFullCheckAsync(CancellationToken.None);
                    return "OK";

                case "STATUS":
                    {
                        bool isBusy = _checkLock.CurrentCount == 0;

                        int minutesLeft = 0;

                        if (_settings.CheckIntervalMinutes > 0)
                        {
                            var remainStatus = _nextTimerRun - DateTime.Now;

                            if (remainStatus.TotalSeconds > 0)
                                minutesLeft = (int)Math.Ceiling(remainStatus.TotalMinutes);
                        }

                        lock (_stateLock)
                        {
                            return $"STATE|BUSY={(isBusy ? 1 : 0)}|PROGRESS={_progressPercent}|STAGE={_currentStage}|TIMER={minutesLeft}";
                        }
                    }

                case "GET_PROGRESS":
                    lock (_stateLock)
                        return $"PROGRESS {_progressPercent}|{_currentStage}";

                case "GET_TIMER":
                    {
                        if (_settings.CheckIntervalMinutes <= 0)
                            return "TIMER DISABLED";

                        var remain = _nextTimerRun - DateTime.Now;

                        if (remain.TotalSeconds <= 0)
                            return "TIMER READY";

                        return $"TIMER {(int)Math.Ceiling(remain.TotalMinutes)}";
                    }
               

                case "GET_LOG":
                    return "LOG " + GetLastServerLogLines();

                case "GET_CERTS":
                    try
                    {
                        var list = _db.LoadAll(true);

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

                    var dparts = originalCmd.Split('|');

                    if (dparts.Length < 3)
                        return "ERROR BAD FORMAT\n";

                    if (!int.TryParse(dparts[1], out int did))
                        return "ERROR BAD ID\n";

                    if (!bool.TryParse(dparts[2], out bool deleted))
                        return "ERROR BAD FLAG\n";

                    _db.MarkAsDeleted(did, deleted);

                    _log($"MARK_DELETED OK: ID={did}, Deleted={deleted}");

                    return "OK DELETED UPDATED\n";

                case "UPDATE_NOTE":
                    {
                        try
                        {
                            var parts = originalCmd.Split(new[] { '|' }, 3);

                            if (parts.Length < 3)
                                return "ERROR BAD FORMAT\n";

                            if (!int.TryParse(parts[1], out int nid))
                                return "ERROR BAD ID\n";

                            string base64Note = parts[2];

                            string note;
                            try
                            {
                                note = Encoding.UTF8.GetString(
                                    Convert.FromBase64String(base64Note));
                            }
                            catch
                            {
                                return "ERROR BAD DATA\n";
                            }

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
                            var aparts = originalCmd.Split(new[] { '|' }, 4);

                            if (aparts.Length < 4)
                                return "ERROR BAD FORMAT\n";

                            if (!int.TryParse(aparts[1], out int certId))
                                return "ERROR BAD ID\n";

                            string fileName = aparts[2];

                            byte[] data;
                            try
                            {
                                data = Convert.FromBase64String(aparts[3]);
                            }
                            catch
                            {
                                return "ERROR BAD DATA\n";
                            }

                            _log($"ADD_ARCHIVE received: ID={certId}, Cert={fileName}, Size={data.Length} bytes");

                            

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
                        var gparts = originalCmd.Split('|');

                        if (gparts.Length < 2)
                            return "ARCHIVE EMPTY\n";

                        if (!int.TryParse(gparts[1], out int gid))
                            return "ARCHIVE EMPTY\n";

                        var result = _db.GetArchiveFromDb(gid);

                        if (result.data == null || result.data.Length == 0)
                            return "ARCHIVE EMPTY\n";

                        _log($"GET_ARCHIVE request: ID={gid}");

                        string fileName = result.fileName;
                        var base64 = Convert.ToBase64String(result.data);

                        _log($"GET_ARCHIVE OK: ID={gid}, Size={result.data.Length}");

                        return "ARCHIVE|" + fileName + "|" + base64;
                    }

                case "SET_BUILDING":
                    {
                        try
                        {
                            var bparts = originalCmd.Split('|');

                            if (bparts.Length < 3)
                                return "ERROR BAD FORMAT\n";

                            if (!int.TryParse(bparts[1], out int bid))
                                return "ERROR BAD ID\n";

                            string building = bparts[2] ?? "";

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
                case "RESET_REVOKES":
                    {
                        try
                        {
                            _log("RESET_REVOKES started");

                            _db.ResetRevocations();

                            _log("RESET_REVOKES finished");

                            return "OK RESET DONE\n";
                        }
                        catch (Exception ex)
                        {
                            _log("RESET_REVOKES ERROR: " + ex.Message);
                            return "ERROR RESET\n";
                        }
                    }
                case "GET_TOKENS":
                    {
                        var tokens = _db.LoadTokens();
                        return "TOKENS " + JsonConvert.SerializeObject(tokens);
                    }
                case "GET_FREE_TOKENS":
                    {
                        var tokens = _db.LoadFreeTokens();
                        return "TOKENS " + JsonConvert.SerializeObject(tokens);
                    }
                case "SET_TOKEN":
                    {
                        var parts = originalCmd.Split('|');

                        if (parts.Length < 3)
                            return "ERROR BAD FORMAT";

                        if (!int.TryParse(parts[1], out int certId))
                            return "ERROR BAD CERT";

                        if (!int.TryParse(parts[2], out int tokenId))
                            return "ERROR BAD TOKEN";

                        _db.AssignToken(tokenId, certId);

                        return "OK";
                    }
                case "ADD_TOKEN":
                    {
                        try
                        {
                            var parts = originalCmd.Split('|');

                            if (parts.Length < 2)
                                return "ERROR|INVALID_SN";

                            var sn = parts[1]?.Trim().ToUpper();

                            if (string.IsNullOrWhiteSpace(sn))
                                return "ERROR|EMPTY_SN";

                            _db.AddToken(sn);

                            return "OK";
                        }
                        catch (FirebirdSql.Data.FirebirdClient.FbException ex)
                        {
                            // 335544665 = unique constraint violation
                            if (ex.ErrorCode == 335544665)
                                return "ERROR|TOKEN_ALREADY_EXISTS";

                            _log("ADD_TOKEN FB ERROR: " + ex.Message);
                            return "ERROR|DB_ERROR";
                        }
                        catch (Exception ex)
                        {
                            _log("ADD_TOKEN ERROR: " + ex.Message);
                            return "ERROR|SERVER_ERROR";
                        }
                    }

                case "DELETE_TOKEN":
                    {
                        var parts = originalCmd.Split('|');

                        if (parts.Length < 2)
                            return "ERROR";

                        if (!int.TryParse(parts[1], out int id))
                            return "ERROR";
                        _db.DeleteToken(id);
                        return "OK";
                    }
                case "UNASSIGN_TOKEN":
                    {
                        var parts = originalCmd.Split('|');

                        if (parts.Length < 2)
                            return "ERROR BAD FORMAT";

                        if (!int.TryParse(parts[1], out int tokenId))
                            return "ERROR BAD TOKEN";

                        _db.UnassignToken(tokenId);

                        return "OK";
                    }




                default:
                    _log("UNKNOWN CMD: " + originalCmd);
                    return "UNKNOWN COMMAND\n";

            }
        }


        public async Task<string> RequestFastCheckAsync(CancellationToken token)
        {
            // защита от параллельного запуска
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                // защита частоты (ПОСЛЕ захвата lock)
                var sinceLast = DateTime.Now - _lastCheckTime;

                if (sinceLast < MinCheckInterval)
                {
                    var wait = MinCheckInterval - sinceLast;

                    _log($"FAST CHECK skipped. Wait {wait.TotalMinutes:F1} min");

                    return $"BUSY WAIT {(int)Math.Ceiling(wait.TotalMinutes)} MIN";
                }


                

                _log("FAST CHECK started");

                lock (_stateLock)
                {
                    _progressPercent = 5;
                    _currentStage = "Connecting";
                }

                var startedAt = DateTime.UtcNow;

                lock (_stateLock)
                {
                    _progressPercent = 30;
                    _currentStage = "Checking mail";
                }

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5)))
                using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.WhenAll(
                        _newWatcher.CheckNewCertificatesFastAsync(linkedCts.Token),
                        _revokeWatcher.CheckRevocationsFastAsync(linkedCts.Token)
                    );
                }

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

                _lastCheckTime = DateTime.Now;

                if (_settings.CheckIntervalMinutes > 0)
                    _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);

                return "OK FAST DONE";
            }
            finally
            {
                
                _checkLock.Release();
            }
        }

        public async Task<string> RequestFullCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Full check started";
                }

                // просто вызываем синхронные методы
                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(15)))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    try
                    {
                        _newWatcher.ProcessNewCertificates(true, linked.Token);
                        _revokeWatcher.ProcessRevocations(true, linked.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        _log("FULL CHECK timeout");
                    }
                }

                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

                if (_settings.CheckIntervalMinutes > 0)
                    _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);

                return "OK FULL DONE";
            }
            finally
            {
                _checkLock.Release();
            }
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
            try
            {
                _tcpServer?.Stop();
            }
            catch { }

            try
            {
                _timer.Dispose();
            }
            catch { }
        }
    }
}

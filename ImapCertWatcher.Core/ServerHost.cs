using ImapCertWatcher.Data;
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

            _tcpServer = new TcpCommandServer(_settings.ServerPort, HandleRemoteCommand);
            _tcpServer.Start();

            _log($"TCP сервер запущен на порту {_settings.ServerPort}");
        }

        private string HandleRemoteCommand(string cmd)
        {
            if (string.IsNullOrWhiteSpace(cmd))
                return "ERROR|EMPTY_COMMAND";

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
                case "FAST_CERTS_CHECK":
                    {
                        var task = RequestFastCertificatesCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

                case "FAST_REVOKES_CHECK":
                    {
                        var task = RequestFastRevocationsCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

                case "FULL_CERTS_CHECK":
                    {
                        var task = RequestFullCertificatesCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

                case "FULL_REVOKES_CHECK":
                    {
                        var task = RequestFullRevocationsCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

                case "FAST_CHECK":
                    {
                        var task = RequestFastCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

                case "FULL_CHECK":
                    {
                        var task = RequestFullCheckAsync(CancellationToken.None);

                        if (task.IsCompleted)
                            return task.GetAwaiter().GetResult();

                        return "OK";
                    }

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

                        return "CERTS " + json;
                    }
                    catch (Exception ex)
                    {
                        _log("GET_CERTS error: " + ex.Message);
                        return "ERROR|GET_CERTS|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                    }

                case "MARK_DELETED":
                    {
                        try
                        {
                            var dparts = originalCmd.Split('|');

                            if (dparts.Length < 3)
                                return "ERROR|MARK_DELETED|BAD_FORMAT";

                            if (!int.TryParse(dparts[1], out int did))
                                return "ERROR|MARK_DELETED|BAD_ID";

                            if (!bool.TryParse(dparts[2], out bool deleted))
                                return "ERROR|MARK_DELETED|BAD_FLAG";

                            _db.MarkAsDeleted(did, deleted);

                            _log($"MARK_DELETED OK: ID={did}, Deleted={deleted}");

                            return "OK DELETED UPDATED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("MARK_DELETED error: " + ex.Message);
                            return "ERROR|MARK_DELETED|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "UPDATE_NOTE":
                    {
                        try
                        {
                            var parts = originalCmd.Split(new[] { '|' }, 3);

                            if (parts.Length < 3)
                                return "ERROR|UPDATE_NOTE|BAD_FORMAT";

                            if (!int.TryParse(parts[1], out int nid))
                                return "ERROR|UPDATE_NOTE|BAD_ID";

                            string base64Note = parts[2];

                            string note;
                            try
                            {
                                note = Encoding.UTF8.GetString(
                                    Convert.FromBase64String(base64Note));
                            }
                            catch
                            {
                                return "ERROR|UPDATE_NOTE|BAD_DATA";
                            }

                            _log($"UPDATE_NOTE received: ID={nid}");

                            _db.UpdateNote(nid, note);

                            _log($"UPDATE_NOTE OK: ID={nid}");

                            return "OK NOTE UPDATED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("UPDATE_NOTE ERROR: " + ex.Message);
                            return "ERROR|UPDATE_NOTE|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "ADD_ARCHIVE":
                    {
                        try
                        {
                            var aparts = originalCmd.Split(new[] { '|' }, 4);

                            if (aparts.Length < 4)
                                return "ERROR|ADD_ARCHIVE|BAD_FORMAT";

                            if (!int.TryParse(aparts[1], out int certId))
                                return "ERROR|ADD_ARCHIVE|BAD_ID";

                            string fileName = aparts[2];

                            if (string.IsNullOrWhiteSpace(fileName))
                                return "ERROR|ADD_ARCHIVE|BAD_FILE_NAME";

                            byte[] data;
                            try
                            {
                                data = Convert.FromBase64String(aparts[3]);
                            }
                            catch
                            {
                                return "ERROR|ADD_ARCHIVE|BAD_DATA";
                            }

                            _log($"ADD_ARCHIVE received: ID={certId}, Cert={fileName}, Size={data.Length} bytes");

                            bool saved = _db.SaveArchiveToDb(certId, fileName, data);

                            if (!saved)
                            {
                                _log($"ADD_ARCHIVE ERROR: SaveArchiveToDb returned false. ID={certId}");
                                return "ERROR|ADD_ARCHIVE|SAVE_FAILED";
                            }

                            _log($"ADD_ARCHIVE OK: ID={certId}");

                            return "OK ARCHIVE SAVED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("ADD_ARCHIVE ERROR: " + ex.Message);
                            return "ERROR|ADD_ARCHIVE|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
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
                                return "ERROR|SET_BUILDING|BAD_FORMAT";

                            if (!int.TryParse(bparts[1], out int bid))
                                return "ERROR|SET_BUILDING|BAD_ID";

                            string building = bparts[2] ?? "";

                            _log($"SET_BUILDING received: ID={bid}, Building='{building}'");

                            _db.UpdateBuilding(bid, building);

                            _log($"SET_BUILDING OK: ID={bid}");

                            return "OK BUILDING UPDATED\n";
                        }
                        catch (Exception ex)
                        {
                            _log("SET_BUILDING ERROR: " + ex.Message);
                            return "ERROR|SET_BUILDING|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
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
                            return "ERROR|RESET_REVOKES|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "GET_TOKENS":
                    {
                        try
                        {
                            var tokens = _db.LoadTokens();
                            return "TOKENS " + JsonConvert.SerializeObject(tokens);
                        }
                        catch (Exception ex)
                        {
                            _log("GET_TOKENS ERROR: " + ex.Message);
                            return "ERROR|GET_TOKENS|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "GET_FREE_TOKENS":
                    {
                        try
                        {
                            var tokens = _db.LoadFreeTokens();
                            return "TOKENS " + JsonConvert.SerializeObject(tokens);
                        }
                        catch (Exception ex)
                        {
                            _log("GET_FREE_TOKENS ERROR: " + ex.Message);
                            return "ERROR|GET_FREE_TOKENS|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "SET_TOKEN":
                    {
                        try
                        {
                            var parts = originalCmd.Split('|');

                            if (parts.Length < 3)
                                return "ERROR|SET_TOKEN|BAD_FORMAT";

                            if (!int.TryParse(parts[1], out int certId))
                                return "ERROR|SET_TOKEN|BAD_CERT";

                            if (!int.TryParse(parts[2], out int tokenId))
                                return "ERROR|SET_TOKEN|BAD_TOKEN";

                            _db.AssignToken(tokenId, certId);

                            return "OK";
                        }
                        catch (Exception ex)
                        {
                            _log("SET_TOKEN ERROR: " + ex.Message);
                            return "ERROR|SET_TOKEN|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "ADD_TOKEN":
                    {
                        try
                        {
                            var parts = originalCmd.Split('|');

                            if (parts.Length < 2)
                                return "ERROR|ADD_TOKEN|INVALID_SN";

                            var sn = parts[1]?.Trim().ToUpperInvariant();

                            if (string.IsNullOrWhiteSpace(sn))
                                return "ERROR|ADD_TOKEN|EMPTY_SN";

                            _db.AddToken(sn);

                            return "OK";
                        }
                        catch (FirebirdSql.Data.FirebirdClient.FbException ex)
                        {
                            // 335544665 = unique constraint violation
                            if (ex.ErrorCode == 335544665)
                                return "ERROR|ADD_TOKEN|TOKEN_ALREADY_EXISTS";

                            _log("ADD_TOKEN FB ERROR: " + ex.Message);
                            return "ERROR|ADD_TOKEN|DB_ERROR";
                        }
                        catch (Exception ex)
                        {
                            _log("ADD_TOKEN ERROR: " + ex.Message);
                            return "ERROR|ADD_TOKEN|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "DELETE_TOKEN":
                    {
                        try
                        {
                            var parts = originalCmd.Split('|');

                            if (parts.Length < 2)
                                return "ERROR|DELETE_TOKEN|BAD_FORMAT";

                            if (!int.TryParse(parts[1], out int id))
                                return "ERROR|DELETE_TOKEN|BAD_ID";

                            _db.DeleteToken(id);

                            return "OK";
                        }
                        catch (Exception ex)
                        {
                            _log("DELETE_TOKEN ERROR: " + ex.Message);
                            return "ERROR|DELETE_TOKEN|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                case "UNASSIGN_TOKEN":
                    {
                        try
                        {
                            var parts = originalCmd.Split('|');

                            if (parts.Length < 2)
                                return "ERROR|UNASSIGN_TOKEN|BAD_FORMAT";

                            if (!int.TryParse(parts[1], out int tokenId))
                                return "ERROR|UNASSIGN_TOKEN|BAD_TOKEN";

                            _db.UnassignToken(tokenId);

                            return "OK";
                        }
                        catch (Exception ex)
                        {
                            _log("UNASSIGN_TOKEN ERROR: " + ex.Message);
                            return "ERROR|UNASSIGN_TOKEN|" + ex.Message.Replace("\r", " ").Replace("\n", " ");
                        }
                    }

                default:
                    _log("UNKNOWN CMD: " + originalCmd);
                    return "UNKNOWN COMMAND\n";

            }
        }


        public async Task<string> RequestFastCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
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
                        Task.Run(() => _newWatcher.ProcessNewCertificates(false, linkedCts.Token), linkedCts.Token),
                        Task.Run(() => _revokeWatcher.ProcessRevocations(false, linkedCts.Token), linkedCts.Token)
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

                _lastCheckTime = DateTime.Now;

                if (_settings.CheckIntervalMinutes > 0)
                    _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);

                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

                _log("FAST CHECK finished");

                return "OK FAST DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST CHECK cancelled");
                return "CANCELLED FAST CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST CHECK ERROR: " + ex.Message);
                return "ERROR FAST CHECK";
            }
            finally
            {
                _checkLock.Release();
            }
        }

        public async Task<string> RequestFastCertificatesCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                _log("FAST CERTS CHECK started");

                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Checking new certificates";
                }

                var startedAt = DateTime.UtcNow;

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5)))
                using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.Run(() => _newWatcher.ProcessNewCertificates(false, linkedCts.Token), linkedCts.Token);
                }

                var newCerts = _db.GetCertificatesAddedAfter(startedAt);

                lock (_stateLock)
                {
                    _progressPercent = 80;
                    _currentStage = "Saving data";
                }

                if (newCerts.Count > 0)
                    _notifications.NotifyNewUsers(newCerts);

                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }


                _log("FAST CERTS CHECK finished");

                return "OK FAST CERTS DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST CERTS CHECK cancelled");
                return "CANCELLED FAST CERTS CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST CERTS CHECK ERROR: " + ex.Message);
                return "ERROR FAST CERTS CHECK";
            }
            finally
            {
                _checkLock.Release();
            }
        }

        public async Task<string> RequestFastRevocationsCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                _log("FAST REVOKES CHECK started");

                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Checking new revocations";
                }

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(5)))
                using (var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.Run(() => _revokeWatcher.ProcessRevocations(false, linkedCts.Token), linkedCts.Token);
                }

                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }


                _log("FAST REVOKES CHECK finished");

                return "OK FAST REVOKES DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST REVOKES CHECK cancelled");
                return "CANCELLED FAST REVOKES CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FAST REVOKES CHECK ERROR: " + ex.Message);
                return "ERROR FAST REVOKES CHECK";
            }
            finally
            {
                _checkLock.Release();
            }
        }

        public async Task<string> RequestFullCertificatesCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                var startedAt = DateTime.UtcNow;

                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Full certificates check";
                }

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(15)))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.Run(() => _newWatcher.ProcessNewCertificates(true, linked.Token), linked.Token);
                }

                var newCerts = _db.GetCertificatesAddedAfter(startedAt);

                lock (_stateLock)
                {
                    _progressPercent = 85;
                    _currentStage = "Saving data";
                }

                if (newCerts.Count > 0)
                    _notifications.NotifyNewUsers(newCerts);


                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

                _log("FULL CERTS CHECK finished");

                return "OK FULL CERTS DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL CERTS CHECK cancelled");
                return "CANCELLED FULL CERTS CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL CERTS CHECK ERROR: " + ex.Message);
                return "ERROR FULL CERTS CHECK";
            }
            finally
            {
                _checkLock.Release();
            }
        }

        public async Task<string> RequestFullRevocationsCheckAsync(CancellationToken token)
        {
            if (!await _checkLock.WaitAsync(0))
                return "BUSY RUNNING";

            try
            {
                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Full revocations check";
                }

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(15)))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.Run(() => _revokeWatcher.ProcessRevocations(true, linked.Token), linked.Token);
                }


                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

                _log("FULL REVOKES CHECK finished");

                return "OK FULL REVOKES DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL REVOKES CHECK cancelled");
                return "CANCELLED FULL REVOKES CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL REVOKES CHECK ERROR: " + ex.Message);
                return "ERROR FULL REVOKES CHECK";
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
                var startedAt = DateTime.UtcNow;

                lock (_stateLock)
                {
                    _progressPercent = 10;
                    _currentStage = "Full check started";
                }

                using (var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(15)))
                using (var linked = CancellationTokenSource.CreateLinkedTokenSource(token, timeoutCts.Token))
                {
                    await Task.Run(() => _newWatcher.ProcessNewCertificates(true, linked.Token), linked.Token);

                    lock (_stateLock)
                    {
                        _progressPercent = 60;
                        _currentStage = "Checking revocations";
                    }

                    await Task.Run(() => _revokeWatcher.ProcessRevocations(true, linked.Token), linked.Token);
                }

                var newCerts = _db.GetCertificatesAddedAfter(startedAt);

                lock (_stateLock)
                {
                    _progressPercent = 85;
                    _currentStage = "Saving data";
                }

                if (newCerts.Count > 0)
                    _notifications.NotifyNewUsers(newCerts);

                _lastCheckTime = DateTime.Now;

                if (_settings.CheckIntervalMinutes > 0)
                    _nextTimerRun = DateTime.Now.AddMinutes(_settings.CheckIntervalMinutes);

                lock (_stateLock)
                {
                    _progressPercent = 100;
                    _currentStage = "IDLE";
                }

                _log("FULL CHECK finished");

                return "OK FULL DONE";
            }
            catch (OperationCanceledException)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL CHECK cancelled");
                return "CANCELLED FULL CHECK";
            }
            catch (Exception ex)
            {
                lock (_stateLock)
                {
                    _progressPercent = 0;
                    _currentStage = "IDLE";
                }

                _log("FULL CHECK ERROR: " + ex.Message);
                return "ERROR FULL CHECK";
            }
            finally
            {
                _checkLock.Release();
            }
        }
        private string GetLastServerLogLines(int count = 300)
        {
            try
            {
                var logFile = LogSession.SessionLogFile;

                if (!File.Exists(logFile))
                    return "";

                var lastLines = File.ReadLines(logFile, Encoding.UTF8)
                    .Reverse()
                    .Take(count)
                    .Reverse();

                return string.Join(Environment.NewLine, lastLines);
            }
            catch (Exception ex)
            {
                _log("GET_LOG ERROR: " + ex.Message);
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

            try
            {
                _checkLock.Dispose();
            }
            catch { }
        }
    }
}

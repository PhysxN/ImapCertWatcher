using FirebirdSql.Data.FirebirdClient;
using ImapCertWatcher.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ImapCertWatcher.Utils;
using System.Text.RegularExpressions;

namespace ImapCertWatcher.Data
{
    public class DbHelper
    {
        private readonly ServerSettings _settings;
        private readonly string _connectionString;
        private readonly Action<string> _addToMiniLog;

        public DbHelper(ServerSettings settings, Action<string> addToMiniLog = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _addToMiniLog = addToMiniLog;
            NormalizeFirebirdPath();
            _connectionString = BuildConnectionString();
            
            Log("DbHelper инициализирован");
        }

        private string BuildConnectionString()
        {
            if (string.IsNullOrWhiteSpace(_settings.FirebirdDbPath))
                throw new InvalidOperationException("Не задан путь к базе Firebird (FirebirdDbPath).");

            if (string.IsNullOrWhiteSpace(_settings.FbServer))
                throw new InvalidOperationException("Не задан сервер Firebird (FbServer).");

            if (string.IsNullOrWhiteSpace(_settings.FbUser))
                throw new InvalidOperationException("Не задан пользователь Firebird (FbUser).");

            if (string.IsNullOrWhiteSpace(_settings.FbPassword))
                throw new InvalidOperationException("Не задан пароль Firebird (FbPassword).");

            var csb = new FbConnectionStringBuilder
            {
                Database = _settings.FirebirdDbPath,
                UserID = _settings.FbUser,
                Password = _settings.FbPassword,
                DataSource = _settings.FbServer,
                Port = 3050,
                Charset = string.IsNullOrWhiteSpace(_settings.FbCharset) ? "UTF8" : _settings.FbCharset,
                Dialect = _settings.FbDialect,
                ServerType = FbServerType.Default,
                Pooling = true,
                MaxPoolSize = 50,
                MinPoolSize = 5
            };

            return csb.ToString();
        }

        #region Initialization: tables, sequences, indices

        private bool IsLocalFirebirdServer()
        {
            var server = (_settings.FbServer ?? "").Trim();

            if (string.IsNullOrWhiteSpace(server))
                return false;

            return server.Equals("127.0.0.1", StringComparison.OrdinalIgnoreCase)
                || server.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                || server.Equals(".", StringComparison.OrdinalIgnoreCase)
                || server.Equals(Environment.MachineName, StringComparison.OrdinalIgnoreCase);
        }

        private void NormalizeFirebirdPath()
        {
            if (string.IsNullOrWhiteSpace(_settings.FirebirdDbPath))
                throw new InvalidOperationException("FirebirdDbPath не задан.");

            // Для удалённого Firebird локальные папки на машине приложения
            // создавать нельзя. Путь должен быть уже корректным путём
            // на стороне Firebird-сервера.
            if (!IsLocalFirebirdServer())
            {
                if (!Path.IsPathRooted(_settings.FirebirdDbPath))
                {
                    throw new InvalidOperationException(
                        "Для удалённого сервера Firebird путь к базе должен быть абсолютным.");
                }

                return;
            }

            // Для локального Firebird допускаем относительный путь рядом с exe
            if (!Path.IsPathRooted(_settings.FirebirdDbPath))
            {
                _settings.FirebirdDbPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    _settings.FirebirdDbPath);
            }

            var dir = Path.GetDirectoryName(_settings.FirebirdDbPath);

            if (string.IsNullOrWhiteSpace(dir))
                throw new InvalidOperationException("Некорректный путь FirebirdDbPath.");

            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        public void EnsureDatabaseAndTable()
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    EnsureSequences(conn);

                    if (!CheckTableExists(conn, "CERTS"))
                    {
                        CreateCertsAndArchivesTables(conn);
                    }
                    else
                    {
                        if (!CheckTableExists(conn, "CERT_ARCHIVES"))
                            CreateArchivesTableOnly(conn);

                        AddMissingColumns(conn);
                    }

                    EnsureIndexesAndConstraints(conn);

                    if (!CheckTableExists(conn, "PROCESSED_MAILS"))
                        CreateProcessedMailsTable(conn);
                    else
                        CreateIndexIfNotExists(conn, "UQ_PROCESSED_MAILS", "PROCESSED_MAILS", "FOLDER_PATH, MAIL_UID, KIND", true);

                    if (!CheckTableExists(conn, "IMAP_LAST_UID"))
                        CreateImapLastUidTable(conn);

                    if (!CheckTableExists(conn, "TOKENS"))
                        CreateTokensTable(conn);
                    else
                        CreateIndexIfNotExists(conn, "UQ_TOKENS_SN", "TOKENS", "SN", true);
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка EnsureDatabaseAndTable: {ex.Message}");
                throw;
            }
        }

        private void EnsureSequences(FbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_CERTS_ID'";
                var r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_CERTS_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_CERTS_ID создан");
                }

                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_CERT_ARCHIVES_ID'";
                r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_CERT_ARCHIVES_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_CERT_ARCHIVES_ID создан");
                }

                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_PROCESSED_MAILS_ID'";
                r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_PROCESSED_MAILS_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_PROCESSED_MAILS_ID создан");
                }

                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_TOKENS_ID'";
                r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_TOKENS_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_TOKENS_ID создан");
                }
            }
        }

        public bool TryOpenConnection(out Exception error)
        {
            error = null;

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = ex;
                return false;
            }
        }


        public void TestConnection()
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT 1 FROM RDB$DATABASE";
                    cmd.ExecuteScalar();
                }
            }
        }

        public void CreateDatabaseExplicit()
        {
            if (!_settings.IsDevelopment)
                throw new InvalidOperationException(
                    "Создание базы запрещено вне режима разработки");

            if (string.IsNullOrWhiteSpace(_settings.FirebirdDbPath))
                throw new InvalidOperationException(
                    "Не задан путь к базе Firebird (FirebirdDbPath).");

            if (string.IsNullOrWhiteSpace(_settings.FbServer))
                throw new InvalidOperationException(
                    "Не задан сервер Firebird (FbServer).");

            if (string.IsNullOrWhiteSpace(_settings.FbUser))
                throw new InvalidOperationException(
                    "Не задан пользователь Firebird (FbUser).");

            if (string.IsNullOrWhiteSpace(_settings.FbPassword))
                throw new InvalidOperationException(
                    "Не задан пароль Firebird (FbPassword).");

            if (!IsLocalFirebirdServer())
            {
                throw new InvalidOperationException(
                    "Автоматическое создание БД разрешено только для локального сервера Firebird (127.0.0.1, localhost, ., имя текущего компьютера).");
            }

            NormalizeFirebirdPath();

            var csb = new FbConnectionStringBuilder
            {
                Database = _settings.FirebirdDbPath,
                UserID = _settings.FbUser,
                Password = _settings.FbPassword,
                DataSource = _settings.FbServer,
                Port = 3050,
                Charset = string.IsNullOrWhiteSpace(_settings.FbCharset) ? "UTF8" : _settings.FbCharset,
                Dialect = _settings.FbDialect,
                ServerType = FbServerType.Default
            };

            FbConnection.CreateDatabase(
                csb.ToString(),
                pageSize: 16384,
                forcedWrites: true,
                overwrite: false);

            Log("База Firebird создана вручную");
        }

        private bool CheckTableExists(FbConnection conn, string tableName)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = @name AND RDB$SYSTEM_FLAG = 0";
                cmd.Parameters.AddWithValue("@name", tableName.ToUpperInvariant());
                var r = Convert.ToInt32(cmd.ExecuteScalar());
                return r > 0;
            }
        }

        private void CreateCertsAndArchivesTables(FbConnection conn)
        {
            using (var tran = conn.BeginTransaction())
            {
                try
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tran;

                        cmd.CommandText = @"
CREATE TABLE CERTS (
    ID INTEGER NOT NULL PRIMARY KEY,
    FIO VARCHAR(200),
    DATE_START TIMESTAMP,
    DATE_END TIMESTAMP,
    DAYS_LEFT INTEGER,
    CERT_NUMBER VARCHAR(100),
    FROM_ADDRESS VARCHAR(200),
    IS_DELETED SMALLINT DEFAULT 0,
    MANUAL_DELETED SMALLINT DEFAULT 0,
    IS_REVOKED SMALLINT DEFAULT 0,
    REVOKE_DATE TIMESTAMP,
    NOTE VARCHAR(1000),
    BUILDING VARCHAR(100),
    FOLDER_PATH VARCHAR(300),
    ARCHIVE_PATH VARCHAR(500),
    MESSAGE_DATE TIMESTAMP,
    DISCOVERED_AT TIMESTAMP,
    TOKEN_ID INTEGER
)";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = @"
CREATE TRIGGER TRG_CERTS_BI FOR CERTS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = NEXT VALUE FOR SEQ_CERTS_ID;
END";
                        cmd.ExecuteNonQuery();

                        Log("Таблица CERTS и триггер созданы");
                    }

                    CreateArchivesTableOnly(conn, tran);

                    tran.Commit();
                }
                catch
                {
                    tran.Rollback();
                    throw;
                }
            }

            // Создаём индексы/уникальность после создания таблиц (с проверкой)
            CreateIndexIfNotExists(conn, "IDX_CERTS_NUMBER", "CERTS", "CERT_NUMBER", false);
            CreateIndexIfNotExists(conn, "UQ_CERTS_FIO", "CERTS", "FIO", true);
        }

        private void CreateArchivesTableOnly(FbConnection conn, FbTransaction transaction = null)
        {
            using (var cmd = conn.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;

                cmd.CommandText = @"
CREATE TABLE CERT_ARCHIVES (
    ID INTEGER NOT NULL PRIMARY KEY,
    CERT_ID INTEGER NOT NULL,
    FILE_NAME VARCHAR(255) NOT NULL,
    FILE_DATA BLOB SUB_TYPE 0,
    FILE_SIZE INTEGER,
    UPLOAD_DATE TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT FK_CERT_ARCHIVES_CERTS FOREIGN KEY (CERT_ID) REFERENCES CERTS(ID) ON DELETE CASCADE
)";
                TryExecuteNoThrow(cmd);

                cmd.CommandText = @"
CREATE TRIGGER TRG_CERT_ARCHIVES_BI FOR CERT_ARCHIVES
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = NEXT VALUE FOR SEQ_CERT_ARCHIVES_ID;
END";
                TryExecuteNoThrow(cmd);

                Log("Таблица CERT_ARCHIVES создана (если не существовала)");
            }
        }

        private void CreateProcessedMailsTable(FbConnection conn, FbTransaction transaction = null)
        {
            using (var cmd = conn.CreateCommand())
            {
                if (transaction != null)
                    cmd.Transaction = transaction;

                cmd.CommandText = @"
CREATE TABLE PROCESSED_MAILS (
    ID INTEGER NOT NULL PRIMARY KEY,
    FOLDER_PATH VARCHAR(300),
    MAIL_UID     VARCHAR(50),
    KIND         VARCHAR(20),
    PROCESSED_DATE TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)";
                TryExecuteNoThrow(cmd);

                cmd.CommandText = @"
CREATE TRIGGER TRG_PROCESSED_MAILS_BI FOR PROCESSED_MAILS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = NEXT VALUE FOR SEQ_PROCESSED_MAILS_ID;
END";
                TryExecuteNoThrow(cmd);

                // Уникальный индекс, чтобы не было дублей по (папка, UID, тип)
                cmd.CommandText = @"
CREATE UNIQUE INDEX UQ_PROCESSED_MAILS
ON PROCESSED_MAILS (FOLDER_PATH, MAIL_UID, KIND)";
                TryExecuteNoThrow(cmd);

                Log("Таблица PROCESSED_MAILS создана (если не существовала)");
            }
        }

        private void CreateImapLastUidTable(FbConnection conn, FbTransaction transaction = null)
        {
            using (var cmd = conn.CreateCommand())
            {
                if (transaction != null)
                    cmd.Transaction = transaction;

                cmd.CommandText = @"
CREATE TABLE IMAP_LAST_UID (
    FOLDER_PATH VARCHAR(500) NOT NULL PRIMARY KEY,
    LAST_UID BIGINT
)";
                TryExecuteNoThrow(cmd);

                Log("Таблица IMAP_LAST_UID создана (если не существовала)");
            }
        }

        private void CreateTokensTable(FbConnection conn, FbTransaction transaction = null)
        {
            using (var cmd = conn.CreateCommand())
            {
                if (transaction != null)
                    cmd.Transaction = transaction;

                cmd.CommandText = @"
CREATE TABLE TOKENS (
    ID INTEGER NOT NULL PRIMARY KEY,
    SN VARCHAR(100) NOT NULL
)";
                TryExecuteNoThrow(cmd);

                cmd.CommandText = @"
CREATE TRIGGER TRG_TOKENS_BI FOR TOKENS
ACTIVE BEFORE INSERT POSITION 0
AS
BEGIN
  IF (NEW.ID IS NULL) THEN
    NEW.ID = NEXT VALUE FOR SEQ_TOKENS_ID;
END";
                TryExecuteNoThrow(cmd);

                cmd.CommandText = @"
CREATE UNIQUE INDEX UQ_TOKENS_SN
ON TOKENS (SN)";
                TryExecuteNoThrow(cmd);

                Log("Таблица TOKENS создана (если не существовала)");
            }
        }

        private void EnsureIndexesAndConstraints(FbConnection conn)
        {
            CreateIndexIfNotExists(conn, "IDX_CERTS_NUMBER", "CERTS", "CERT_NUMBER", false);
            CreateIndexIfNotExists(conn, "UQ_CERTS_FIO", "CERTS", "FIO", true);
            CreateIndexIfNotExists(conn, "IDX_CERTS_REVOKED", "CERTS", "IS_REVOKED", false);
            CreateIndexIfNotExists(conn, "IDX_CERTS_FILTER", "CERTS", "IS_DELETED, BUILDING, DATE_END", false);
            CreateIndexIfNotExists(conn, "IDX_CERTS_CERT_DELETED", "CERTS", "CERT_NUMBER, IS_DELETED", false);
            CreateIndexIfNotExists(conn, "IDX_CERTS_DISCOVERED_AT", "CERTS", "DISCOVERED_AT", false);

            // Один токен = один сертификат
            CreateIndexIfNotExists(conn, "UQ_CERTS_TOKEN_ID", "CERTS", "TOKEN_ID", true);
        }
        private void AddMissingColumns(FbConnection conn)
        {
            var needed = new Dictionary<string, string>
{
    {"DAYS_LEFT", "INTEGER"},
    {"CERT_NUMBER", "VARCHAR(100)"},
    {"FROM_ADDRESS", "VARCHAR(200)"},
    {"IS_DELETED", "SMALLINT DEFAULT 0"},
    {"MANUAL_DELETED", "SMALLINT DEFAULT 0"},
    {"NOTE", "VARCHAR(1000)"},
    {"BUILDING", "VARCHAR(100)"},
    {"FOLDER_PATH", "VARCHAR(300)"},
    {"ARCHIVE_PATH", "VARCHAR(500)"},
    {"MESSAGE_DATE", "TIMESTAMP"},
    {"DISCOVERED_AT", "TIMESTAMP"},
    {"IS_REVOKED", "SMALLINT DEFAULT 0"},
    {"REVOKE_DATE", "TIMESTAMP"},
    {"TOKEN_ID", "INTEGER"}
};
            foreach (var kv in needed)
            {
                if (!ColumnExists(conn, kv.Key))
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = $"ALTER TABLE CERTS ADD {kv.Key} {kv.Value}";
                        TryExecuteNoThrow(cmd);
                        Log($"Добавлен столбец {kv.Key}");
                    }
                }
            }
        }

        private bool ColumnExists(FbConnection conn, string columnName)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE RDB$RELATION_NAME = 'CERTS' AND RDB$FIELD_NAME = @col";
                cmd.Parameters.AddWithValue("@col", columnName.ToUpperInvariant());
                var r = Convert.ToInt32(cmd.ExecuteScalar());
                return r > 0;
            }
        }

        private void TryExecuteNoThrow(FbCommand cmd)
        {
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Log($"DDL error: {ex.Message}");
                throw;
            }
        }
        /// <summary>
        /// Проверяет существование индекса по имени (RDB$INDICES)
        /// </summary>
        private bool IndexExists(FbConnection conn, string indexName)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$INDICES WHERE UPPER(RDB$INDEX_NAME) = @name";
                cmd.Parameters.AddWithValue("@name", indexName.ToUpperInvariant());
                var cnt = Convert.ToInt32(cmd.ExecuteScalar());
                return cnt > 0;
            }
        }

        /// <summary>
        /// Создаёт индекс, если он ещё не существует.
        /// </summary>
        private void CreateIndexIfNotExists(FbConnection conn, string indexName, string tableName, string columns, bool unique)
        {
            if (IndexExists(conn, indexName))
                return;

            using (var cmd = conn.CreateCommand())
            {
                string uniq = unique ? "UNIQUE " : "";
                cmd.CommandText = $"CREATE {uniq}INDEX {indexName} ON {tableName} ({columns})";
                cmd.ExecuteNonQuery();
                Log($"Создан индекс {indexName} ON {tableName} ({columns})");
            }
        }
        #endregion

        #region Normalization utilities

        // Нормализуем ФИО для сравнения: убираем лишние пробелы, приводим к верхнему регистру, заменяем ё->е
        public static string NormalizeFio(string fio)
        {
            if (string.IsNullOrWhiteSpace(fio)) return string.Empty;
            var s = fio.Trim();
            s = Regex.Replace(s, @"\s+", " ");
            s = s.Replace('ё', 'е').Replace('Ё', 'Е');
            s = s.ToUpperInvariant();
            return s;
        }

        // Нормализуем номер сертификата: убираем "№", пробелы, приводим к верхнему регистру
        public static string NormalizeCertNumber(string cert)
        {
            if (string.IsNullOrWhiteSpace(cert))
                return string.Empty;

            var s = cert.Trim();

            s = Regex.Replace(
                s,
                @"^\s*(?:СЕРТИФИКАТ\s*)?(?:№|N)\s*",
                "",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            s = Regex.Replace(s, @"[\s\-]+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        public List<CertRecord> GetCertificatesAddedAfter(DateTime fromUtc)
        {
            var list = new List<CertRecord>();

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT
    c.ID,
    c.FIO,
    c.DATE_START,
    c.DATE_END,
    c.CERT_NUMBER,
    c.FROM_ADDRESS,
    c.IS_DELETED,
    c.IS_REVOKED,
    c.REVOKE_DATE,
    c.NOTE,
    c.BUILDING,
    c.FOLDER_PATH,
    c.ARCHIVE_PATH,
    c.MESSAGE_DATE,
    c.TOKEN_ID,
    t.SN AS TOKEN_SN,
    (SELECT COUNT(*) FROM CERT_ARCHIVES ca WHERE ca.CERT_ID = c.ID) AS HAS_ARCHIVE
FROM CERTS c
LEFT JOIN TOKENS t ON t.ID = c.TOKEN_ID
WHERE c.DISCOVERED_AT >= @dt
ORDER BY c.DATE_END";

                        cmd.Parameters.AddWithValue("@dt", fromUtc);

                        using (var rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                list.Add(MapReaderToCertRecord(rdr));
                            }
                        }
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                Log($"GetCertificatesAddedAfter error: {ex.Message}");
                throw;
            }
        }


        /// <summary>
        /// Сохраняет файл архива в таблицу CERT_ARCHIVES (не транзакционно).
        /// </summary>
        public bool SaveArchiveToDb(int certId, string filePath, string fileName = null)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Log($"SaveArchiveToDb: файл не найден {filePath}");
                    return false;
                }

                byte[] data = File.ReadAllBytes(filePath);
                fileName = fileName ?? Path.GetFileName(filePath);
                return SaveArchiveToDb(certId, fileName, data);
            }
            catch (Exception ex)
            {
                Log($"SaveArchiveToDb exception: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Сохраняет файл архива в таблицу CERT_ARCHIVES по байтам (не транзакционно).
        /// </summary>
        public bool SaveArchiveToDb(int certId, string fileName, byte[] fileData)
        {
            try
            {
                if (certId <= 0)
                {
                    Log("SaveArchiveToDb: некорректный certId");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(fileName))
                {
                    Log("SaveArchiveToDb: пустое имя файла");
                    return false;
                }

                if (fileData == null || fileData.Length == 0)
                {
                    Log("SaveArchiveToDb: пустые данные файла");
                    return false;
                }

                fileName = Path.GetFileName(fileName);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var tr = conn.BeginTransaction())
                    {
                        try
                        {
                            using (var del = conn.CreateCommand())
                            {
                                del.Transaction = tr;
                                del.CommandText = "DELETE FROM CERT_ARCHIVES WHERE CERT_ID = @certId";
                                del.Parameters.AddWithValue("@certId", certId);
                                del.ExecuteNonQuery();
                            }

                            using (var ins = conn.CreateCommand())
                            {
                                ins.Transaction = tr;
                                ins.CommandText = @"
INSERT INTO CERT_ARCHIVES (CERT_ID, FILE_NAME, FILE_DATA, FILE_SIZE, UPLOAD_DATE)
VALUES (@certId, @fileName, @fileData, @fileSize, @uploadDate)";
                                ins.Parameters.AddWithValue("@certId", certId);
                                ins.Parameters.AddWithValue("@fileName", fileName);
                                ins.Parameters.AddWithValue("@fileData", fileData);
                                ins.Parameters.AddWithValue("@fileSize", fileData.Length);
                                ins.Parameters.AddWithValue("@uploadDate", DateTime.Now);
                                ins.ExecuteNonQuery();
                            }

                            using (var upd = conn.CreateCommand())
                            {
                                upd.Transaction = tr;
                                upd.CommandText = @"
UPDATE CERTS
SET ARCHIVE_PATH = @archivePath
WHERE ID = @id";
                                upd.Parameters.AddWithValue("@archivePath", fileName);
                                upd.Parameters.AddWithValue("@id", certId);
                                upd.ExecuteNonQuery();
                            }

                            tr.Commit();
                        }
                        catch
                        {
                            try { tr.Rollback(); } catch { }
                            throw;
                        }
                    }
                }

                Log($"SaveArchiveToDb: сохранён архив '{fileName}' для CertID={certId}");
                return true;
            }
            catch (Exception ex)
            {
                Log($"SaveArchiveToDb error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Транзакционное сохранение архива и опциональное обновление поля ARCHIVE_PATH в CERTS.
        /// </summary>
        public bool SaveArchiveToDbTransactional(int certId, string filePath, string fileName = null)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Log($"SaveArchiveToDbTransactional: файл не найден: {filePath}");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = Path.GetFileName(filePath);

                var fileData = File.ReadAllBytes(filePath);

                return SaveArchiveToDb(certId, fileName, fileData);
            }
            catch (Exception ex)
            {
                Log($"SaveArchiveToDbTransactional: исключение: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Возвращает последний архив (data, fileName) для certId.
        /// </summary>
        public (byte[] data, string fileName) GetArchiveFromDb(int certId)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT FILE_DATA, FILE_NAME
FROM CERT_ARCHIVES
WHERE CERT_ID = @certId
ORDER BY ID DESC
ROWS 1";
                        cmd.Parameters.AddWithValue("@certId", certId);

                        using (var rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                var blob = (byte[])rdr["FILE_DATA"];
                                var fname = rdr["FILE_NAME"].ToString();

                                Log($"GetArchiveFromDb: найден архив {fname} для CertID={certId}");
                                return (blob, fname);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetArchiveFromDb error: {ex.Message}");
            }

            return (null, null);
        }

        public bool HasArchiveInDb(int certId)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM CERT_ARCHIVES WHERE CERT_ID = @certId";
                        cmd.Parameters.AddWithValue("@certId", certId);
                        var r = Convert.ToInt32(cmd.ExecuteScalar());
                        return r > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"HasArchiveInDb error: {ex.Message}");
                return false;
            }
        }

        public bool DeleteArchiveFromDb(int certId)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var tr = conn.BeginTransaction())
                    {
                        try
                        {
                            int deletedRows;

                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tr;
                                cmd.CommandText = "DELETE FROM CERT_ARCHIVES WHERE CERT_ID = @certId";
                                cmd.Parameters.AddWithValue("@certId", certId);
                                deletedRows = cmd.ExecuteNonQuery();
                            }

                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tr;
                                cmd.CommandText = @"
UPDATE CERTS
SET ARCHIVE_PATH = NULL
WHERE ID = @certId";
                                cmd.Parameters.AddWithValue("@certId", certId);
                                cmd.ExecuteNonQuery();
                            }

                            tr.Commit();

                            Log($"DeleteArchiveFromDb: удалено архивов={deletedRows} для CertID={certId}");
                            return deletedRows > 0;
                        }
                        catch
                        {
                            try { tr.Rollback(); } catch { }
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"DeleteArchiveFromDb error: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Find / duplicate checks / get by id





        #endregion

        #region InsertOrUpdate (atomic variant included)

        /// <summary>
        /// Старый интерфейс совместимости: возвращает (wasUpdated, wasAdded)
        /// </summary>
        public (bool wasUpdated, bool wasAdded) InsertOrUpdate(CertEntry entry)
        {
            var (wasUpdated, wasAdded, _) = InsertOrUpdateAndGetId(entry);
            return (wasUpdated, wasAdded);
        }

        /// <summary>
        /// Атомарный InsertOrUpdate: возвращает (wasUpdated, wasAdded, certId).
        /// Логика:
        /// - одно ФИО = одна запись;
        /// - если запись по ФИО уже есть:
        ///   - если пришёл тот же сертификат или более новый/не старее — обновляем запись;
        ///   - если пришёл более старый сертификат — ничего не меняем;
        /// - если записи по ФИО нет — создаём новую.
        /// Вся операция происходит в транзакции.
        /// </summary>
        /// 
        public (bool wasUpdated, bool wasAdded, int certId) InsertOrUpdateAndGetId(CertEntry entry)
        {
            if (entry == null)
                throw new ArgumentNullException(nameof(entry));

            bool wasUpdated = false;
            bool wasAdded = false;
            int certId = -1;

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var tx = conn.BeginTransaction())
                    {
                        try
                        {
                            string fioKey = NormalizeFio(entry.Fio);
                            string newCertNorm = NormalizeCertNumber(entry.CertNumber);

                            int existingId = -1;
                            DateTime existingDateEnd = DateTime.MinValue;
                            string existingCertNumber = null;
                            string existingFromAddress = null;
                            bool existingIsDeleted = false;
                            bool existingIsRevoked = false;

                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = @"
SELECT ID, FIO, DATE_END, CERT_NUMBER, FROM_ADDRESS, IS_DELETED, IS_REVOKED
FROM CERTS
ORDER BY DATE_END DESC";

                                using (var r = cmd.ExecuteReader())
                                {
                                    while (r.Read())
                                    {
                                        string dbFio = r.IsDBNull(1) ? "" : r.GetString(1);

                                        if (NormalizeFio(dbFio) != fioKey)
                                            continue;

                                        existingId = r.GetInt32(0);
                                        existingDateEnd = r.IsDBNull(2) ? DateTime.MinValue : r.GetDateTime(2);
                                        existingCertNumber = r.IsDBNull(3) ? "" : r.GetString(3);
                                        existingFromAddress = r.IsDBNull(4) ? "" : r.GetString(4);
                                        existingIsDeleted = !r.IsDBNull(5) && Convert.ToInt32(r[5]) == 1;
                                        existingIsRevoked = !r.IsDBNull(6) && Convert.ToInt32(r[6]) == 1;
                                        break;
                                    }
                                }
                            }

                            if (existingId > 0)
                            {
                                string existingCertNorm = NormalizeCertNumber(existingCertNumber);

                                bool sameCert = existingCertNorm == newCertNorm;
                                bool sameOrNewer = entry.DateEnd >= existingDateEnd;

                                certId = existingId;

                                if (sameCert)
                                {
                                    if (existingIsRevoked || existingIsDeleted)
                                    {
                                        Log($"InsertOrUpdateAndGetId: найден тот же сертификат для FIO='{entry.Fio}', номер={entry.CertNumber}, " +
                                            $"но запись уже имеет статус удалённой/аннулированной. Повторное чтение старого письма не восстанавливает запись.");
                                    }
                                    else
                                    {
                                        Log($"InsertOrUpdateAndGetId: тот же сертификат для FIO='{entry.Fio}', номер={entry.CertNumber} уже есть в БД. Запись не изменена.");
                                    }
                                }
                                else if (sameOrNewer)
                                {
                                    string fromAddressToSave = string.IsNullOrWhiteSpace(entry.FromAddress)
                                        ? (existingFromAddress ?? "")
                                        : entry.FromAddress;

                                    using (var upd = conn.CreateCommand())
                                    {
                                        upd.Transaction = tx;
                                        upd.CommandText = @"
UPDATE CERTS SET
    FIO             = @fioOrig,
    CERT_NUMBER     = @certOrig,
    DATE_START      = @dateStart,
    DATE_END        = @dateEnd,
    DAYS_LEFT       = 0,
    FROM_ADDRESS    = @fromAddress,
    FOLDER_PATH     = @folderPath,
    MESSAGE_DATE    = @msgDate,
    DISCOVERED_AT   = @discoveredAt,
    IS_DELETED      = 0,
    MANUAL_DELETED  = 0,
    IS_REVOKED      = 0,
    REVOKE_DATE     = NULL
WHERE ID = @id";
                                        upd.Parameters.AddWithValue("@fioOrig", (object)entry.Fio ?? "");
                                        upd.Parameters.AddWithValue("@certOrig", (object)entry.CertNumber ?? "");
                                        upd.Parameters.AddWithValue("@dateStart",
                                            entry.DateStart == DateTime.MinValue ? (object)DBNull.Value : entry.DateStart);
                                        upd.Parameters.AddWithValue("@dateEnd",
                                            entry.DateEnd == DateTime.MinValue ? (object)DBNull.Value : entry.DateEnd);
                                        upd.Parameters.AddWithValue("@fromAddress", fromAddressToSave);
                                        upd.Parameters.AddWithValue("@folderPath", (object)entry.FolderPath ?? "");
                                        upd.Parameters.AddWithValue("@msgDate",
                                            entry.MessageDate == DateTime.MinValue ? (object)DBNull.Value : entry.MessageDate);
                                        upd.Parameters.AddWithValue("@discoveredAt", DateTime.UtcNow);
                                        upd.Parameters.AddWithValue("@id", existingId);

                                        int affected = upd.ExecuteNonQuery();

                                        if (affected <= 0)
                                            throw new InvalidOperationException("Не удалось обновить существующий сертификат.");
                                    }

                                    wasUpdated = true;

                                    if (existingIsDeleted || existingIsRevoked)
                                    {
                                        Log($"InsertOrUpdateAndGetId: запись ID={existingId} переведена в активное состояние новым сертификатом");
                                    }

                                    Log($"InsertOrUpdateAndGetId: замена сертификата для FIO='{entry.Fio}': " +
                                        $"старый номер={existingCertNumber}, новый={entry.CertNumber}, " +
                                        $"старый DATE_END={existingDateEnd:dd.MM.yyyy HH:mm:ss}, " +
                                        $"новый DATE_END={entry.DateEnd:dd.MM.yyyy HH:mm:ss}");
                                }
                                else
                                {
                                    Log($"InsertOrUpdateAndGetId: найден более старый сертификат для FIO='{entry.Fio}', " +
                                        $"номер={entry.CertNumber}, DATE_END={entry.DateEnd:dd.MM.yyyy HH:mm:ss}. " +
                                        $"В БД уже есть более новый до {existingDateEnd:dd.MM.yyyy HH:mm:ss}. Запись не изменена.");
                                }
                            }
                            else
                            {
                                using (var ins = conn.CreateCommand())
                                {
                                    ins.Transaction = tx;
                                    ins.CommandText = @"
INSERT INTO CERTS (
    FIO, CERT_NUMBER, DATE_START, DATE_END, DAYS_LEFT,
    FROM_ADDRESS, FOLDER_PATH, MESSAGE_DATE, DISCOVERED_AT
) VALUES (
    @fioOrig, @certOrig, @dateStart, @dateEnd, @daysLeft,
    @fromAddress, @folderPath, @msgDate, @discoveredAt
)
RETURNING ID";
                                    ins.Parameters.AddWithValue("@fioOrig", (object)entry.Fio ?? "");
                                    ins.Parameters.AddWithValue("@certOrig", (object)entry.CertNumber ?? "");
                                    ins.Parameters.AddWithValue("@dateStart",
                                        entry.DateStart == DateTime.MinValue ? (object)DBNull.Value : entry.DateStart);
                                    ins.Parameters.AddWithValue("@dateEnd",
                                        entry.DateEnd == DateTime.MinValue ? (object)DBNull.Value : entry.DateEnd);
                                    ins.Parameters.AddWithValue("@daysLeft", 0);
                                    ins.Parameters.AddWithValue("@fromAddress", (object)entry.FromAddress ?? "");
                                    ins.Parameters.AddWithValue("@folderPath", (object)entry.FolderPath ?? "");
                                    ins.Parameters.AddWithValue("@msgDate",
                                        entry.MessageDate == DateTime.MinValue ? (object)DBNull.Value : entry.MessageDate);
                                    ins.Parameters.AddWithValue("@discoveredAt", DateTime.UtcNow);

                                    var insertedIdObj = ins.ExecuteScalar();

                                    if (insertedIdObj == null || insertedIdObj == DBNull.Value)
                                        throw new InvalidOperationException("Не удалось получить ID новой записи сертификата.");

                                    certId = Convert.ToInt32(insertedIdObj);
                                    wasAdded = true;

                                    Log($"InsertOrUpdateAndGetId: добавлен новый сертификат ID={certId} для FIO='{entry.Fio}', номер={entry.CertNumber}");
                                }
                            }

                            tx.Commit();
                        }
                        catch
                        {
                            try { tx.Rollback(); } catch { }
                            throw;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                try { Log($"InsertOrUpdateAndGetId error: {ex.Message}"); } catch { }
                throw;
            }

            return (wasUpdated, wasAdded, certId);
        }

        private CertRecord MapReaderToCertRecord(FbDataReader rdr)
        {
            var rec = new CertRecord();

            rec.Id = rdr["ID"] == DBNull.Value ? 0 : Convert.ToInt32(rdr["ID"]);
            rec.Fio = rdr["FIO"]?.ToString();
            rec.CertNumber = rdr["CERT_NUMBER"]?.ToString();
            rec.FromAddress = rdr["FROM_ADDRESS"]?.ToString();
            rec.Note = rdr["NOTE"]?.ToString();
            rec.Building = rdr["BUILDING"]?.ToString();
            rec.FolderPath = rdr["FOLDER_PATH"]?.ToString();
            rec.ArchivePath = rdr["ARCHIVE_PATH"]?.ToString();

            rec.DateStart = rdr["DATE_START"] == DBNull.Value
                ? DateTime.MinValue
                : (DateTime)rdr["DATE_START"];

            rec.DateEnd = rdr["DATE_END"] == DBNull.Value
                ? DateTime.MinValue
                : (DateTime)rdr["DATE_END"];

            rec.MessageDate = rdr["MESSAGE_DATE"] == DBNull.Value
                ? DateTime.MinValue
                : (DateTime)rdr["MESSAGE_DATE"];

            // Firebird SMALLINT -> bool
            rec.IsDeleted = rdr["IS_DELETED"] != DBNull.Value &&
                            Convert.ToInt32(rdr["IS_DELETED"]) == 1;
            rec.IsRevoked = rdr["IS_REVOKED"] != DBNull.Value &&
                Convert.ToInt32(rdr["IS_REVOKED"]) == 1;

            rec.RevokeDate = rdr["REVOKE_DATE"] == DBNull.Value
                ? (DateTime?)null
                : (DateTime)rdr["REVOKE_DATE"];
            rec.TokenId = rdr["TOKEN_ID"] == DBNull.Value
                ? (int?)null
                : Convert.ToInt32(rdr["TOKEN_ID"]);
            rec.TokenSn = rdr["TOKEN_SN"] == DBNull.Value
                ? null
                : rdr["TOKEN_SN"].ToString();

            // Подзапрос HAS_ARCHIVE
            if (rdr["HAS_ARCHIVE"] != DBNull.Value)
                rec.HasArchive = Convert.ToInt32(rdr["HAS_ARCHIVE"]) > 0;

            return rec;
        }


        public long GetLastUid(string folderPath)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
SELECT LAST_UID
FROM IMAP_LAST_UID
WHERE FOLDER_PATH = @p";

                    cmd.Parameters.AddWithValue("@p", folderPath);

                    var val = cmd.ExecuteScalar();
                    return val == null || val == DBNull.Value ? 0 : Convert.ToInt64(val);
                }
            }
        }

        public void UpdateLastUid(string folderPath, long lastUid)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
UPDATE OR INSERT INTO IMAP_LAST_UID (FOLDER_PATH, LAST_UID)
VALUES (@p, @u)
MATCHING (FOLDER_PATH)";

                    cmd.Parameters.AddWithValue("@p", folderPath);
                    cmd.Parameters.AddWithValue("@u", lastUid);

                    cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Загружает все записи с фильтрами.
        /// </summary>
        public List<CertRecord> LoadAll(bool showDeleted = false, string buildingFilter = "Все")
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            var list = new List<CertRecord>();
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    var where = new List<string>();


                    if (!showDeleted)
                    {
                        where.Add("(c.IS_DELETED IS NULL OR c.IS_DELETED = 0)");
                    }

                    if (!string.IsNullOrEmpty(buildingFilter) && buildingFilter != "Все")
                    {
                        where.Add("c.BUILDING = @building");
                        cmd.Parameters.AddWithValue("@building", buildingFilter);
                    }

                    string whereClause = where.Count > 0
                        ? "WHERE " + string.Join(" AND ", where)
                        : "";

                    cmd.CommandText = $@"
                                        SELECT
                                            c.ID,
                                            c.FIO,
                                            c.DATE_START,
                                            c.DATE_END,
                                            c.CERT_NUMBER,
                                            c.FROM_ADDRESS,
                                            c.IS_DELETED,
                                            c.IS_REVOKED,
                                            c.REVOKE_DATE,
                                            c.NOTE,
                                            c.BUILDING,
                                            c.FOLDER_PATH,
                                            c.ARCHIVE_PATH,
                                            c.MESSAGE_DATE,
                                            c.TOKEN_ID,
                                            t.SN AS TOKEN_SN,
                                            (SELECT COUNT(*) FROM CERT_ARCHIVES ca WHERE ca.CERT_ID = c.ID) AS HAS_ARCHIVE
                                        FROM CERTS c
                                        LEFT JOIN TOKENS t ON t.ID = c.TOKEN_ID
                                        {whereClause}
                                        ORDER BY c.DATE_END";

                    using (var rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(MapReaderToCertRecord(rdr));
                        }
                    }
                }
            }

            sw.Stop();  // ⬅ стоп секундомера


            // Log($"LoadAll: загружено {list.Count} записей (showDeleted={showDeleted}, filter={buildingFilter})");
            return list;
        }

        public void UpdateNote(int id, string note)
        {
            var newNote = note ?? "";
            CertLogInfo before;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    before = LoadCertLogInfo(conn, tr, id);
                    if (before == null)
                        throw new InvalidOperationException("Сертификат не найден.");

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = "UPDATE CERTS SET NOTE = @note WHERE ID = @id";
                        cmd.Parameters.AddWithValue("@note", newNote);
                        cmd.Parameters.AddWithValue("@id", id);

                        int affected = cmd.ExecuteNonQuery();

                        if (affected <= 0)
                            throw new InvalidOperationException("Сертификат не найден.");
                    }

                    tr.Commit();
                }
            }

            Log(
                $"UpdateNote: сертификат ID={id}, ФИО='{LogValue(before.Fio)}'. " +
                $"NOTE: '{LogValue(before.Note)}' -> '{LogValue(newNote)}'.");
        }

        public void UpdateBuilding(int id, string building)
        {
            var newBuilding = building ?? "";
            CertLogInfo before;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    before = LoadCertLogInfo(conn, tr, id);
                    if (before == null)
                        throw new InvalidOperationException("Сертификат не найден.");

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = "UPDATE CERTS SET BUILDING = @b WHERE ID = @id";
                        cmd.Parameters.AddWithValue("@b", newBuilding);
                        cmd.Parameters.AddWithValue("@id", id);

                        int affected = cmd.ExecuteNonQuery();

                        if (affected <= 0)
                            throw new InvalidOperationException("Сертификат не найден.");
                    }

                    tr.Commit();
                }
            }

            Log(
                $"UpdateBuilding: сертификат ID={id}, ФИО='{LogValue(before.Fio)}'. " +
                $"BUILDING: '{LogValue(before.Building)}' -> '{LogValue(newBuilding)}'.");
        }

        public void MarkAsDeleted(int id, bool isDeleted)
        {
            CertLogInfo before;
            CertLogInfo after;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    try
                    {
                        before = LoadCertLogInfo(conn, tr, id);
                        if (before == null)
                            throw new InvalidOperationException("Сертификат не найден.");

                        if (!isDeleted && before.IsRevoked)
                        {
                            throw new InvalidOperationException(
                                "Нельзя восстанавливать аннулированный сертификат обычным восстановлением. Используйте сброс аннулирований.");
                        }

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tr;

                            if (isDeleted)
                            {
                                cmd.CommandText = @"
UPDATE CERTS
SET
    IS_DELETED = 1,
    MANUAL_DELETED = 1
WHERE ID = @id";
                            }
                            else
                            {
                                cmd.CommandText = @"
UPDATE CERTS
SET
    IS_DELETED = 0,
    MANUAL_DELETED = 0
WHERE ID = @id
  AND COALESCE(IS_REVOKED, 0) = 0";
                            }

                            cmd.Parameters.AddWithValue("@id", id);

                            int affected = cmd.ExecuteNonQuery();

                            if (affected <= 0)
                                throw new InvalidOperationException("Не удалось изменить состояние записи.");
                        }

                        after = LoadCertLogInfo(conn, tr, id);

                        tr.Commit();
                    }
                    catch
                    {
                        try { tr.Rollback(); } catch { }
                        throw;
                    }
                }
            }

            Log(
                $"MarkAsDeleted: сертификат ID={id}, ФИО='{LogValue(before.Fio)}'. " +
                $"IS_DELETED: {LogBool(before.IsDeleted)} -> {LogBool(after != null && after.IsDeleted)}. " +
                $"IS_REVOKED={LogBool(after != null && after.IsRevoked)}.");
        }

        public void UpdateArchivePath(int id, string archivePath)
        {
            var newArchivePath = archivePath ?? "";
            CertLogInfo before;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    before = LoadCertLogInfo(conn, tr, id);
                    if (before == null)
                        throw new InvalidOperationException("Сертификат не найден.");

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = "UPDATE CERTS SET ARCHIVE_PATH = @p WHERE ID = @id";
                        cmd.Parameters.AddWithValue("@p", newArchivePath);
                        cmd.Parameters.AddWithValue("@id", id);

                        int affected = cmd.ExecuteNonQuery();

                        if (affected <= 0)
                            throw new InvalidOperationException("Сертификат не найден.");
                    }

                    tr.Commit();
                }
            }

            Log(
                $"UpdateArchivePath: сертификат ID={id}, ФИО='{LogValue(before.Fio)}'. " +
                $"ARCHIVE_PATH: '{LogValue(before.ArchivePath)}' -> '{LogValue(newArchivePath)}'.");
        }

        public FbConnection GetConnection() => new FbConnection(_connectionString);

        #endregion

        #region Processed mails helpers

        /// <summary>
        /// Проверяет, обрабатывали ли мы уже письмо с указанным UID в папке и отмеченным типом.
        /// kind: например, "NEW" (новые сертификаты) или "REVOKE" (аннулирования).
        /// </summary>
        public bool IsMailProcessed(string folderPath, string mailUid, string kind)
        {
            if (string.IsNullOrEmpty(mailUid)) return false;
            folderPath = folderPath ?? "";
            kind = string.IsNullOrEmpty(kind) ? "GENERIC" : kind.ToUpperInvariant();

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT COUNT(*) 
FROM PROCESSED_MAILS 
WHERE FOLDER_PATH = @folder AND MAIL_UID = @uid AND KIND = @kind";
                        cmd.Parameters.AddWithValue("@folder", folderPath);
                        cmd.Parameters.AddWithValue("@uid", mailUid);
                        cmd.Parameters.AddWithValue("@kind", kind);

                        var r = Convert.ToInt32(cmd.ExecuteScalar());
                        return r > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"IsMailProcessed error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Помечает письмо как обработанное (если уже есть запись — молча игнорируем ошибку уникальности).
        /// </summary>
        public void MarkMailProcessed(string folderPath, string mailUid, string kind)
        {
            if (string.IsNullOrEmpty(mailUid))
                return;

            folderPath = folderPath ?? "";
            kind = string.IsNullOrEmpty(kind) ? "GENERIC" : kind.ToUpperInvariant();

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
UPDATE OR INSERT INTO PROCESSED_MAILS
(FOLDER_PATH, MAIL_UID, KIND, PROCESSED_DATE)
VALUES (@folder, @uid, @kind, @dt)
MATCHING (FOLDER_PATH, MAIL_UID, KIND)";

                        cmd.Parameters.AddWithValue("@folder", folderPath);
                        cmd.Parameters.AddWithValue("@uid", mailUid);
                        cmd.Parameters.AddWithValue("@kind", kind);
                        cmd.Parameters.AddWithValue("@dt", DateTime.Now);

                        cmd.ExecuteNonQuery();
                    }
                }

                // Логируем ОДИН раз после успешного upsert
            }
            catch (Exception ex)
            {
                Log($"MarkMailProcessed ERROR: {ex.Message}");
                throw;
            }
        }

        #endregion

        public List<TokenRecord> LoadTokens()
        {
            var list = new List<TokenRecord>();

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                                        SELECT
                                            t.ID,
                                            t.SN,
                                            c.ID AS OWNER_CERT_ID,
                                            c.FIO
                                        FROM TOKENS t
                                        LEFT JOIN CERTS c ON c.TOKEN_ID = t.ID
                                        ORDER BY t.SN";

                    using (var rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(new TokenRecord
                            {
                                Id = rdr.GetInt32(0),
                                Sn = rdr.GetString(1),
                                OwnerCertId = rdr.IsDBNull(2) ? (int?)null : rdr.GetInt32(2),
                                OwnerFio = rdr.IsDBNull(3) ? null : rdr.GetString(3)
                            });
                        }
                    }
                }
            }

            return list;
        }

        public List<TokenRecord> LoadFreeTokens()
        {
            var list = new List<TokenRecord>();

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
                                        SELECT t.ID, t.SN
                                        FROM TOKENS t
                                        WHERE NOT EXISTS (
                                            SELECT 1 FROM CERTS c WHERE c.TOKEN_ID = t.ID
                                        )
                                        ORDER BY t.SN";

                    using (var rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            list.Add(new TokenRecord
                            {
                                Id = rdr.GetInt32(0),
                                Sn = rdr.GetString(1),
                                OwnerCertId = null
                            });
                        }
                    }
                }
            }

            return list;
        }
        public void AssignToken(int tokenId, int certId)
        {
            string tokenSn = null;
            string certFio = null;
            int? previousOwnerCertId = null;
            string previousOwnerFio = null;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    try
                    {
                        using (var checkCert = conn.CreateCommand())
                        {
                            checkCert.Transaction = tr;
                            checkCert.CommandText = @"
SELECT
    FIO,
    COALESCE(IS_DELETED, 0),
    COALESCE(IS_REVOKED, 0)
FROM CERTS
WHERE ID = @certId";
                            checkCert.Parameters.AddWithValue("@certId", certId);

                            using (var rdr = checkCert.ExecuteReader())
                            {
                                if (!rdr.Read())
                                    throw new InvalidOperationException("Сертификат не найден.");

                                certFio = rdr.IsDBNull(0) ? null : rdr.GetString(0);

                                bool isDeleted = Convert.ToInt32(rdr[1]) == 1;
                                bool isRevoked = Convert.ToInt32(rdr[2]) == 1;

                                if (isDeleted)
                                    throw new InvalidOperationException("Нельзя назначить токен удалённому сертификату.");

                                if (isRevoked)
                                    throw new InvalidOperationException("Нельзя назначить токен аннулированному сертификату.");
                            }
                        }

                        using (var checkToken = conn.CreateCommand())
                        {
                            checkToken.Transaction = tr;
                            checkToken.CommandText = @"
SELECT SN
FROM TOKENS
WHERE ID = @tokenId";
                            checkToken.Parameters.AddWithValue("@tokenId", tokenId);

                            var result = checkToken.ExecuteScalar();
                            if (result == null || result == DBNull.Value)
                                throw new InvalidOperationException("Токен не найден.");

                            tokenSn = Convert.ToString(result);
                        }

                        using (var oldOwnerCmd = conn.CreateCommand())
                        {
                            oldOwnerCmd.Transaction = tr;
                            oldOwnerCmd.CommandText = @"
SELECT FIRST 1
    ID,
    FIO
FROM CERTS
WHERE TOKEN_ID = @tokenId";
                            oldOwnerCmd.Parameters.AddWithValue("@tokenId", tokenId);

                            using (var rdr = oldOwnerCmd.ExecuteReader())
                            {
                                if (rdr.Read())
                                {
                                    previousOwnerCertId = rdr.IsDBNull(0) ? (int?)null : rdr.GetInt32(0);
                                    previousOwnerFio = rdr.IsDBNull(1) ? null : rdr.GetString(1);
                                }
                            }
                        }

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tr;
                            cmd.CommandText = @"
UPDATE CERTS
SET TOKEN_ID = NULL
WHERE TOKEN_ID = @tokenId";
                            cmd.Parameters.AddWithValue("@tokenId", tokenId);
                            cmd.ExecuteNonQuery();
                        }

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tr;
                            cmd.CommandText = @"
UPDATE CERTS
SET TOKEN_ID = @tokenId
WHERE ID = @certId";
                            cmd.Parameters.AddWithValue("@tokenId", tokenId);
                            cmd.Parameters.AddWithValue("@certId", certId);

                            int affected = cmd.ExecuteNonQuery();
                            if (affected <= 0)
                                throw new InvalidOperationException("Не удалось назначить токен сертификату.");
                        }

                        tr.Commit();
                    }
                    catch
                    {
                        try { tr.Rollback(); } catch { }
                        throw;
                    }
                }
            }

            if (previousOwnerCertId.HasValue && previousOwnerCertId.Value != certId)
            {
                Log(
                    $"AssignToken: токен '{LogValue(tokenSn)}' (ID={tokenId}) назначен сертификату ID={certId}, ФИО='{LogValue(certFio)}'. " +
                    $"Ранее токен был привязан к сертификату ID={previousOwnerCertId.Value}, ФИО='{LogValue(previousOwnerFio)}'.");
            }
            else
            {
                Log(
                    $"AssignToken: токен '{LogValue(tokenSn)}' (ID={tokenId}) назначен сертификату ID={certId}, ФИО='{LogValue(certFio)}'.");
            }
        }

        public void AddToken(string sn)
        {
            string normalizedSn = (sn ?? "").Trim().ToUpperInvariant();
            int? newTokenId = null;

            if (string.IsNullOrWhiteSpace(normalizedSn))
                throw new InvalidOperationException("Серийный номер токена пустой.");

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = @"
INSERT INTO TOKENS (SN)
VALUES (@sn)";
                        cmd.Parameters.AddWithValue("@sn", normalizedSn);
                        cmd.ExecuteNonQuery();
                    }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = @"
SELECT FIRST 1 ID
FROM TOKENS
WHERE SN = @sn
ORDER BY ID DESC";
                        cmd.Parameters.AddWithValue("@sn", normalizedSn);

                        var value = cmd.ExecuteScalar();
                        if (value != null && value != DBNull.Value)
                            newTokenId = Convert.ToInt32(value);
                    }

                    tr.Commit();
                }
            }

            Log($"AddToken: токен '{normalizedSn}' (ID={newTokenId?.ToString() ?? "?"}) добавлен.");
        }

        public void DeleteToken(int id)
        {
            TokenLogInfo tokenInfo;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    tokenInfo = LoadTokenLogInfo(conn, tr, id);
                    if (tokenInfo == null)
                        throw new InvalidOperationException("Токен не найден.");

                    using (var checkUsage = conn.CreateCommand())
                    {
                        checkUsage.Transaction = tr;
                        checkUsage.CommandText = "SELECT COUNT(*) FROM CERTS WHERE TOKEN_ID = @id";
                        checkUsage.Parameters.AddWithValue("@id", id);

                        var inUse = Convert.ToInt32(checkUsage.ExecuteScalar()) > 0;
                        if (inUse)
                            throw new InvalidOperationException("Нельзя удалить занятый токен.");
                    }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = @"DELETE FROM TOKENS WHERE ID = @id";
                        cmd.Parameters.AddWithValue("@id", id);

                        int affected = cmd.ExecuteNonQuery();
                        if (affected <= 0)
                            throw new InvalidOperationException("Не удалось удалить токен.");
                    }

                    tr.Commit();
                }
            }

            Log($"DeleteToken: токен '{LogValue(tokenInfo.Sn)}' (ID={id}) удалён.");
        }

        public void UnassignToken(int tokenId)
        {
            string tokenSn = null;
            int? ownerCertId = null;
            string ownerFio = null;

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                {
                    using (var checkToken = conn.CreateCommand())
                    {
                        checkToken.Transaction = tr;
                        checkToken.CommandText = @"
SELECT SN
FROM TOKENS
WHERE ID = @id";
                        checkToken.Parameters.AddWithValue("@id", tokenId);

                        var result = checkToken.ExecuteScalar();
                        if (result == null || result == DBNull.Value)
                            throw new InvalidOperationException("Токен не найден.");

                        tokenSn = Convert.ToString(result);
                    }

                    using (var ownerCmd = conn.CreateCommand())
                    {
                        ownerCmd.Transaction = tr;
                        ownerCmd.CommandText = @"
SELECT FIRST 1
    ID,
    FIO
FROM CERTS
WHERE TOKEN_ID = @id";
                        ownerCmd.Parameters.AddWithValue("@id", tokenId);

                        using (var rdr = ownerCmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                ownerCertId = rdr.IsDBNull(0) ? (int?)null : rdr.GetInt32(0);
                                ownerFio = rdr.IsDBNull(1) ? null : rdr.GetString(1);
                            }
                        }
                    }

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tr;
                        cmd.CommandText = @"
UPDATE CERTS
SET TOKEN_ID = NULL
WHERE TOKEN_ID = @id";
                        cmd.Parameters.AddWithValue("@id", tokenId);
                        cmd.ExecuteNonQuery();
                    }

                    tr.Commit();
                }
            }

            if (ownerCertId.HasValue)
            {
                Log(
                    $"UnassignToken: токен '{LogValue(tokenSn)}' (ID={tokenId}) снят с сертификата ID={ownerCertId.Value}, ФИО='{LogValue(ownerFio)}'.");
            }
            else
            {
                Log(
                    $"UnassignToken: токен '{LogValue(tokenSn)}' (ID={tokenId}) был свободен, привязки не было.");
            }
        }

        public void ResetRevocations()
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var tr = conn.BeginTransaction())
                using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = tr;

                    cmd.CommandText = @"
UPDATE CERTS
SET
    IS_REVOKED = 0,
    IS_DELETED = CASE
        WHEN COALESCE(MANUAL_DELETED, 0) = 1 THEN 1
        ELSE 0
    END,
    REVOKE_DATE = NULL
WHERE IS_REVOKED = 1";

                    int certs = cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"
DELETE FROM PROCESSED_MAILS
WHERE KIND = 'REVOKE'";

                    int mails = cmd.ExecuteNonQuery();

                    cmd.Parameters.Clear();
                    cmd.CommandText = @"
DELETE FROM IMAP_LAST_UID
WHERE RIGHT(FOLDER_PATH, 7) = '_REVOKE'";

                    int uids = cmd.ExecuteNonQuery();

                    tr.Commit();

                    Log($"ResetRevocations: CERTS={certs}, MAILS={mails}, IMAP_LAST_UID={uids}");
                }
            }
        }

        /// <summary>
        /// Находит сертификат по номеру и помечает как аннулированный, добавляя в NOTE дату отзыва (если указана).
        /// Возвращает true если найден и обновлён.
        /// </summary>

        public bool FindAndMarkAsRevokedByCertAndFio(
            string certNumber,
            string fio,
            string folderPath,
            string revokeDateShort,
            out string resultMessage)
        {
            resultMessage = "ошибка при обработке аннулирования";

            Log($"FindAndMarkAsRevokedByCertAndFio: cert={certNumber}, fio={fio}, date={revokeDateShort}");

            try
            {
                var certNorm = NormalizeCertNumber(certNumber);
                var fioNorm = NormalizeFio(fio);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var tx = conn.BeginTransaction())
                    {
                        var candidates = new List<(int Id, string CertNumber, string Fio, int IsRevoked, int IsDeleted)>();

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = @"
SELECT
    ID,
    CERT_NUMBER,
    FIO,
    COALESCE(IS_REVOKED, 0) AS IS_REVOKED,
    COALESCE(IS_DELETED, 0) AS IS_DELETED
FROM CERTS
WHERE CERT_NUMBER IS NOT NULL
ORDER BY ID DESC";

                            using (var rdr = cmd.ExecuteReader())
                            {
                                while (rdr.Read())
                                {
                                    string dbCertNumber = rdr["CERT_NUMBER"] == DBNull.Value
                                        ? ""
                                        : rdr["CERT_NUMBER"].ToString();

                                    if (NormalizeCertNumber(dbCertNumber) != certNorm)
                                        continue;

                                    candidates.Add((
                                        Convert.ToInt32(rdr["ID"]),
                                        dbCertNumber,
                                        rdr["FIO"] == DBNull.Value ? "" : rdr["FIO"].ToString(),
                                        Convert.ToInt32(rdr["IS_REVOKED"]),
                                        Convert.ToInt32(rdr["IS_DELETED"])
                                    ));
                                }
                            }
                        }

                        if (candidates.Count == 0)
                        {
                            resultMessage = "номер не найден в текущей БД, вероятно сертификат уже заменён";
                            Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. certNorm={certNorm}");
                            return false;
                        }

                        (int Id, string CertNumber, string Fio, int IsRevoked, int IsDeleted)? selected = null;

                        if (!string.IsNullOrWhiteSpace(fioNorm))
                        {
                            var fioMatches = candidates
                                .Where(x => NormalizeFio(x.Fio) == fioNorm)
                                .ToList();

                            if (fioMatches.Count == 0)
                            {
                                var dbFios = string.Join(" | ", candidates.Select(x => x.Fio).Distinct().Take(5));
                                resultMessage = "номер найден, но ФИО не совпало с текущей БД";
                                Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. certNorm={certNorm}, fioNorm={fioNorm}, dbFios={dbFios}");
                                return false;
                            }

                            selected = fioMatches
                                .OrderBy(x => x.IsRevoked)
                                .ThenBy(x => x.IsDeleted)
                                .ThenByDescending(x => x.Id)
                                .First();
                        }
                        else
                        {
                            if (candidates.Count == 1)
                            {
                                selected = candidates[0];
                                Log($"FindAndMarkAsRevokedByCertAndFio: ФИО не извлечено, но по номеру найден ровно один кандидат. ID={selected.Value.Id}");
                            }
                            else
                            {
                                var activeCandidates = candidates
                                    .Where(x => x.IsRevoked == 0)
                                    .OrderBy(x => x.IsDeleted)
                                    .ThenByDescending(x => x.Id)
                                    .ToList();

                                if (activeCandidates.Count == 1)
                                {
                                    selected = activeCandidates[0];
                                    Log($"FindAndMarkAsRevokedByCertAndFio: ФИО не извлечено, выбран единственный активный кандидат. ID={selected.Value.Id}");
                                }
                                else
                                {
                                    resultMessage = "ФИО не извлечено, а по номеру найдено несколько кандидатов";
                                    Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. certNorm={certNorm}, count={candidates.Count}");
                                    return false;
                                }
                            }
                        }

                        var row = selected.Value;

                        if (row.IsRevoked == 1)
                        {
                            resultMessage = "номер найден, запись уже аннулирована";
                            Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. ID={row.Id}, CERT_NUMBER={row.CertNumber}, FIO={row.Fio}");
                            return false;
                        }

                        DateTime parsedDate = DateTime.Now;

                        if (!string.IsNullOrWhiteSpace(revokeDateShort))
                        {
                            DateTime dt;

                            if (DateTime.TryParseExact(
                                    revokeDateShort,
                                    "dd.MM.yyyy HH:mm:ss",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None,
                                    out dt))
                            {
                                parsedDate = dt;
                            }
                            else if (DateTime.TryParseExact(
                                    revokeDateShort,
                                    "dd.MM.yyyy",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None,
                                    out dt))
                            {
                                parsedDate = dt;
                            }
                            else if (DateTime.TryParse(revokeDateShort, out dt))
                            {
                                parsedDate = dt;
                            }
                        }

                        using (var upd = conn.CreateCommand())
                        {
                            upd.Transaction = tx;
                            upd.CommandText = @"
UPDATE CERTS
SET
    IS_DELETED = 1,
    IS_REVOKED = 1,
    REVOKE_DATE = @revDate
WHERE ID = @id
  AND COALESCE(IS_REVOKED, 0) = 0";

                            upd.Parameters.AddWithValue("@revDate", parsedDate);
                            upd.Parameters.AddWithValue("@id", row.Id);

                            int rows = upd.ExecuteNonQuery();

                            if (rows == 0)
                            {
                                resultMessage = "номер найден, но обновление не выполнилось";
                                Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. ID={row.Id}, CERT_NUMBER={row.CertNumber}, FIO={row.Fio}, IS_REVOKED={row.IsRevoked}, IS_DELETED={row.IsDeleted}");
                                return false;
                            }
                        }

                        tx.Commit();

                        resultMessage = "номер найден и аннулирован";
                        Log($"FindAndMarkAsRevokedByCertAndFio: {resultMessage}. ID={row.Id}, CERT_NUMBER={row.CertNumber}, FIO={row.Fio}, revokeDate={parsedDate:dd.MM.yyyy HH:mm:ss}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                resultMessage = "ошибка при обработке аннулирования";
                Log($"FindAndMarkAsRevokedByCertAndFio error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Находит по FIO+certNumber и помечает аннулированным.
        /// </summary>



        #region Logging

        private sealed class CertLogInfo
        {
            public int Id { get; set; }
            public string Fio { get; set; }
            public string Building { get; set; }
            public string Note { get; set; }
            public string ArchivePath { get; set; }
            public bool IsDeleted { get; set; }
            public bool IsRevoked { get; set; }
            public int? TokenId { get; set; }
        }

        private sealed class TokenLogInfo
        {
            public int Id { get; set; }
            public string Sn { get; set; }
        }

        private CertLogInfo LoadCertLogInfo(FbConnection conn, FbTransaction tr, int id)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tr;
                cmd.CommandText = @"
SELECT
    ID,
    FIO,
    BUILDING,
    NOTE,
    ARCHIVE_PATH,
    COALESCE(IS_DELETED, 0),
    COALESCE(IS_REVOKED, 0),
    TOKEN_ID
FROM CERTS
WHERE ID = @id";
                cmd.Parameters.AddWithValue("@id", id);

                using (var rdr = cmd.ExecuteReader())
                {
                    if (!rdr.Read())
                        return null;

                    return new CertLogInfo
                    {
                        Id = rdr.GetInt32(0),
                        Fio = rdr.IsDBNull(1) ? null : rdr.GetString(1),
                        Building = rdr.IsDBNull(2) ? null : rdr.GetString(2),
                        Note = rdr.IsDBNull(3) ? null : rdr.GetString(3),
                        ArchivePath = rdr.IsDBNull(4) ? null : rdr.GetString(4),
                        IsDeleted = Convert.ToInt32(rdr[5]) == 1,
                        IsRevoked = Convert.ToInt32(rdr[6]) == 1,
                        TokenId = rdr.IsDBNull(7) ? (int?)null : rdr.GetInt32(7)
                    };
                }
            }
        }

        private TokenLogInfo LoadTokenLogInfo(FbConnection conn, FbTransaction tr, int id)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.Transaction = tr;
                cmd.CommandText = @"
SELECT ID, SN
FROM TOKENS
WHERE ID = @id";
                cmd.Parameters.AddWithValue("@id", id);

                using (var rdr = cmd.ExecuteReader())
                {
                    if (!rdr.Read())
                        return null;

                    return new TokenLogInfo
                    {
                        Id = rdr.GetInt32(0),
                        Sn = rdr.IsDBNull(1) ? null : rdr.GetString(1)
                    };
                }
            }
        }

        private static string LogValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "<пусто>" : value;
        }

        private static string LogBool(bool value)
        {
            return value ? "1" : "0";
        }

        private void Log(string message)
        {
            try
            {
                string line = $"[DbHelper] {message}";

                if (_addToMiniLog != null)
                    _addToMiniLog(line);
                else
                    System.IO.File.AppendAllText(
                        ImapCertWatcher.Utils.LogSession.SessionLogFile,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {line}{Environment.NewLine}",
                        System.Text.Encoding.UTF8);

                System.Diagnostics.Debug.WriteLine(line);
            }
            catch
            {
                // не фейлим приложение из-за логирования
            }
        }
        #endregion
    }
}

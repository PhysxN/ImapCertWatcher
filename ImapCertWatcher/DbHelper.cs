using FirebirdSql.Data.FirebirdClient;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ImapCertWatcher.Data
{
    public class DbHelper
    {
        private readonly AppSettings _settings;
        private readonly string _connectionString;
        private readonly Action<string> _addToMiniLog;

        public DbHelper(AppSettings settings, Action<string> addToMiniLog = null)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _addToMiniLog = addToMiniLog;
            _connectionString = BuildConnectionString();

            EnsureDatabaseAndTable();
            Log("DbHelper инициализирован");
        }

        private string BuildConnectionString()
        {
            var csb = new FbConnectionStringBuilder
            {
                Database = _settings.FirebirdDbPath,
                UserID = string.IsNullOrEmpty(_settings.FbUser) ? "SYSDBA" : _settings.FbUser,
                Password = string.IsNullOrEmpty(_settings.FbPassword) ? "masterkey" : _settings.FbPassword,
                DataSource = string.IsNullOrEmpty(_settings.FbServer) ? "127.0.0.1" : _settings.FbServer,
                Port = 3050,
                // по умолчанию используем WIN1251 если не указано
                Charset = string.IsNullOrEmpty(_settings.FbCharset) ? "WIN1251" : _settings.FbCharset,
                Dialect = _settings.FbDialect,
                ServerType = FbServerType.Default,
                Pooling = false
            };

            Log($"Построена строка подключения: Server={csb.DataSource}, DB={csb.Database}, Charset={csb.Charset}");
            return csb.ToString();
        }

        #region Initialization: tables, sequences, indices

        public void EnsureDatabaseAndTable()
        {
            try
            {
                // Если локальная БД не существует — попытка создать (только для локального сервера)
                if ((string.IsNullOrEmpty(_settings.FbServer) || _settings.FbServer == "127.0.0.1" || _settings.FbServer == "localhost")
                    && !File.Exists(_settings.FirebirdDbPath))
                {
                    var dir = Path.GetDirectoryName(_settings.FirebirdDbPath);
                    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                    var createCsb = new FbConnectionStringBuilder
                    {
                        Database = _settings.FirebirdDbPath,
                        UserID = string.IsNullOrEmpty(_settings.FbUser) ? "SYSDBA" : _settings.FbUser,
                        Password = string.IsNullOrEmpty(_settings.FbPassword) ? "masterkey" : _settings.FbPassword,
                        DataSource = string.IsNullOrEmpty(_settings.FbServer) ? "127.0.0.1" : _settings.FbServer,
                        Port = 3050,
                        Charset = string.IsNullOrEmpty(_settings.FbCharset) ? "WIN1251" : _settings.FbCharset,
                        Dialect = _settings.FbDialect,
                        ServerType = FbServerType.Default
                    };

                    FbConnection.CreateDatabase(createCsb.ToString(), pageSize: 16384, forcedWrites: true, overwrite: false);
                    Log("Локальная база Firebird создана.");
                }

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    // Создаем последовательности, таблицы и индексы, если необходимо
                    EnsureSequences(conn);
                    if (!CheckTableExists(conn, "CERTS"))
                        CreateCertsAndArchivesTables(conn);
                    else
                    {
                        // Таблица есть — создаём недостающие таблицы/индексы
                        if (!CheckTableExists(conn, "CERT_ARCHIVES"))
                            CreateArchivesTableOnly(conn);

                        EnsureIndexesAndConstraints(conn);
                        AddMissingColumns(conn);
                    }

                    // --- НОВОЕ: таблица PROCESSED_MAILS ---
                    if (!CheckTableExists(conn, "PROCESSED_MAILS"))
                        CreateProcessedMailsTable(conn);

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
                // SEQ_CERTS_ID
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_CERTS_ID'";
                var r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_CERTS_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_CERTS_ID создан");
                }

                // SEQ_CERT_ARCHIVES_ID
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_CERT_ARCHIVES_ID'";
                r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_CERT_ARCHIVES_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_CERT_ARCHIVES_ID создан");
                }

                // --- НОВОЕ: генератор для PROCESSED_MAILS ---
                cmd.CommandText = @"SELECT COUNT(*) FROM RDB$GENERATORS WHERE RDB$GENERATOR_NAME = 'SEQ_PROCESSED_MAILS_ID'";
                r = Convert.ToInt32(cmd.ExecuteScalar());
                if (r == 0)
                {
                    cmd.CommandText = "CREATE SEQUENCE SEQ_PROCESSED_MAILS_ID";
                    cmd.ExecuteNonQuery();
                    Log("SEQ_PROCESSED_MAILS_ID создан");
                }
            }
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
    NOTE VARCHAR(1000),
    BUILDING VARCHAR(100),
    FOLDER_PATH VARCHAR(300),
    ARCHIVE_PATH VARCHAR(500),
    MESSAGE_DATE TIMESTAMP
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
            CreateIndexIfNotExists(conn, "IDX_CERTS_FIO", "CERTS", "FIO", false);
            CreateIndexIfNotExists(conn, "UQ_CERTS_FIO_NUMBER", "CERTS", "FIO, CERT_NUMBER", true);
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

        private void EnsureIndexesAndConstraints(FbConnection conn)
        {
            CreateIndexIfNotExists(conn, "IDX_CERTS_NUMBER", "CERTS", "CERT_NUMBER", false);
            CreateIndexIfNotExists(conn, "IDX_CERTS_FIO", "CERTS", "FIO", false);
            CreateIndexIfNotExists(conn, "UQ_CERTS_FIO_NUMBER", "CERTS", "FIO, CERT_NUMBER", true);

            // новый индекс под фильтры грида
            CreateIndexIfNotExists(conn, "IDX_CERTS_FILTER", "CERTS", "IS_DELETED, BUILDING, DATE_END", false);
        }

        private void AddMissingColumns(FbConnection conn)
        {
            var needed = new Dictionary<string, string>
            {
                {"CERT_NUMBER", "VARCHAR(100)"},
                {"FROM_ADDRESS", "VARCHAR(200)"},
                {"IS_DELETED", "SMALLINT DEFAULT 0"},
                {"NOTE", "VARCHAR(1000)"},
                {"BUILDING", "VARCHAR(100)"},
                {"FOLDER_PATH", "VARCHAR(300)"},
                {"ARCHIVE_PATH", "VARCHAR(500)"},
                {"MESSAGE_DATE", "TIMESTAMP"}
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
                // Игнорируем ошибки создания индексов/триггеров, но логируем
                Log($"TryExecuteNoThrow: {ex.Message}");
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
            try
            {
                if (IndexExists(conn, indexName))
                {
                    Log($"Индекс {indexName} уже существует — пропускаем создание.");
                    return;
                }

                using (var cmd = conn.CreateCommand())
                {
                    string uniq = unique ? "UNIQUE " : "";
                    cmd.CommandText = $"CREATE {uniq}INDEX {indexName} ON {tableName} ({columns})";
                    cmd.ExecuteNonQuery();
                    Log($"Создан индекс {indexName} ON {tableName} ({columns})");
                }
            }
            catch (Exception ex)
            {
                // Если индекс не удалось создать — логируем, но не даём падать и остальному коду
                Log($"CreateIndexIfNotExists: не удалось создать индекс {indexName}: {ex.Message}");
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
            if (string.IsNullOrWhiteSpace(cert)) return string.Empty;
            var s = cert.Trim();
            s = s.Replace("№", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();
            return s;
        }

        #endregion

        #region Archive methods (transactional + non-transactional)

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
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"INSERT INTO CERT_ARCHIVES (CERT_ID, FILE_NAME, FILE_DATA, FILE_SIZE)
                                            VALUES (@certId, @fileName, @fileData, @fileSize)";
                        cmd.Parameters.AddWithValue("@certId", certId);
                        cmd.Parameters.AddWithValue("@fileName", fileName);
                        cmd.Parameters.AddWithValue("@fileData", fileData);
                        cmd.Parameters.AddWithValue("@fileSize", fileData.Length);
                        cmd.ExecuteNonQuery();
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

                if (string.IsNullOrEmpty(fileName))
                    fileName = Path.GetFileName(filePath);

                byte[] fileData = File.ReadAllBytes(filePath);
                long fileSize = fileData.LongLength;

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            // 1) Вставляем строку в CERT_ARCHIVES
                            using (var insertCmd = conn.CreateCommand())
                            {
                                insertCmd.Transaction = transaction;
                                insertCmd.CommandText = @"
INSERT INTO CERT_ARCHIVES (CERT_ID, FILE_NAME, FILE_DATA, FILE_SIZE, UPLOAD_DATE)
VALUES (@certId, @fileName, @fileData, @fileSize, @uploadDate)";
                                insertCmd.Parameters.AddWithValue("@certId", certId);
                                insertCmd.Parameters.AddWithValue("@fileName", fileName);
                                insertCmd.Parameters.AddWithValue("@fileData", fileData);
                                insertCmd.Parameters.AddWithValue("@fileSize", fileSize);
                                insertCmd.Parameters.AddWithValue("@uploadDate", DateTime.Now);

                                insertCmd.ExecuteNonQuery();
                            }

                            // 2) Обновляем поле ARCHIVE_PATH в CERTS
                            using (var updateCmd = conn.CreateCommand())
                            {
                                updateCmd.Transaction = transaction;
                                updateCmd.CommandText = @"
UPDATE CERTS
SET ARCHIVE_PATH = @archivePath
WHERE ID = @id";
                                updateCmd.Parameters.AddWithValue("@archivePath", fileName ?? "");
                                updateCmd.Parameters.AddWithValue("@id", certId);

                                updateCmd.ExecuteNonQuery();
                            }

                            transaction.Commit();
                            Log($"SaveArchiveToDbTransactional: успешно сохранён архив для CertID={certId}, file={fileName} ({fileSize} bytes)");
                            return true;
                        }
                        catch (Exception exTx)
                        {
                            try { transaction.Rollback(); } catch { }
                            Log($"SaveArchiveToDbTransactional: ошибка транзакции: {exTx.Message}");
                            return false;
                        }
                    }
                }
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
SELECT FILE_DATA, FILE_NAME FROM CERT_ARCHIVES
WHERE CERT_ID = @certId
ORDER BY UPLOAD_DATE DESC
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
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "DELETE FROM CERT_ARCHIVES WHERE CERT_ID = @certId";
                        cmd.Parameters.AddWithValue("@certId", certId);
                        int r = cmd.ExecuteNonQuery();
                        Log($"DeleteArchiveFromDb: удалено записей={r} для CertID={certId}");
                        return r > 0;
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

        public int GetCertIdByNumberAndFio(string certNumber, string fio)
        {
            try
            {
                var certNorm = NormalizeCertNumber(certNumber);
                var fioNorm = NormalizeFio(fio);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT ID FROM CERTS WHERE UPPER(TRIM(CERT_NUMBER)) = @certNumber AND UPPER(TRIM(FIO)) = @fio";
                        cmd.Parameters.AddWithValue("@certNumber", certNorm);
                        cmd.Parameters.AddWithValue("@fio", fioNorm);
                        var r = cmd.ExecuteScalar();
                        return r != null ? Convert.ToInt32(r) : -1;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"GetCertIdByNumberAndFio error: {ex.Message}");
                return -1;
            }
        }

        public CertRecord FindByFio(string fio)
        {
            if (string.IsNullOrWhiteSpace(fio)) return null;
            var fioNorm = NormalizeFio(fio);

            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        // Берём самую "свежую" запись по дате окончания
                        cmd.CommandText = @"
SELECT ID, FIO, DATE_START, DATE_END, DAYS_LEFT, CERT_NUMBER, FROM_ADDRESS, IS_DELETED, NOTE, BUILDING, FOLDER_PATH, ARCHIVE_PATH, MESSAGE_DATE
FROM CERTS
WHERE UPPER(TRIM(FIO)) = @fio
ORDER BY DATE_END DESC
ROWS 1";
                        cmd.Parameters.AddWithValue("@fio", fioNorm);
                        using (var rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                return MapReaderToCertRecord(rdr);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"FindByFio error: {ex.Message}");
            }
            return null;
        }

        public bool CheckDuplicate(string fio, string certNumber)
        {
            try
            {
                var fioNorm = NormalizeFio(fio);
                var certNorm = NormalizeCertNumber(certNumber);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT COUNT(*) FROM CERTS WHERE UPPER(TRIM(FIO)) = @fio AND UPPER(TRIM(CERT_NUMBER)) = @cert";
                        cmd.Parameters.AddWithValue("@fio", fioNorm);
                        cmd.Parameters.AddWithValue("@cert", certNorm);
                        var r = Convert.ToInt32(cmd.ExecuteScalar());
                        return r > 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"CheckDuplicate error: {ex.Message}");
                return false;
            }
        }

        public bool IsDuplicateCertificate(string fio, string certNumber) => CheckDuplicate(fio, certNumber);

        #endregion

        #region InsertOrUpdate (atomic variant included)

        /// <summary>
        /// Старый интерфейс совместимости: возвращает (wasUpdated, wasAdded)
        /// </summary>
        public (bool wasUpdated, bool wasAdded) InsertOrUpdate(CertEntry entry)
        {
            var res = InsertOrUpdateAndGetId(entry);
            return (res.wasUpdated, res.wasAdded);
        }

        /// <summary>
        /// Атомарный InsertOrUpdate: возвращает (wasUpdated, wasAdded, certId).
        /// - если запись найдена по FIO + CERT_NUMBER: если сообщение новее - обновляем, иначе пропускаем.
        /// - если не найдена - создаём новую запись.
        /// Вся операция происходит в транзакции.
        /// </summary>
        public (bool wasUpdated, bool wasAdded, int certId) InsertOrUpdateAndGetId(CertEntry entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));

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
                            string fioKey = (entry.Fio ?? "").Trim();
                            string certKey = (entry.CertNumber ?? "").Trim();

                            int existingId = -1;
                            DateTime existingDateEnd = DateTime.MinValue;
                            string existingCertNumber = null;

                            // 1) Попытка найти по FIO + CERT_NUMBER
                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = tx;
                                cmd.CommandText = @"
SELECT ID, DATE_END, CERT_NUMBER
FROM CERTS
WHERE UPPER(TRIM(FIO)) = UPPER(TRIM(@fio))
  AND UPPER(TRIM(CERT_NUMBER)) = UPPER(TRIM(@cert))
ROWS 1";
                                cmd.Parameters.AddWithValue("@fio", fioKey);
                                cmd.Parameters.AddWithValue("@cert", certKey);

                                using (var r = cmd.ExecuteReader())
                                {
                                    if (r.Read())
                                    {
                                        existingId = r.GetInt32(0);
                                        existingDateEnd = r.IsDBNull(1) ? DateTime.MinValue : r.GetDateTime(1);
                                        existingCertNumber = r.IsDBNull(2) ? "" : r.GetString(2);
                                    }
                                }
                            }

                            // 2) Если не нашли по FIO+CERT_NUMBER — ищем по одному ФИО
                            if (existingId <= 0)
                            {
                                using (var cmd = conn.CreateCommand())
                                {
                                    cmd.Transaction = tx;
                                    cmd.CommandText = @"
SELECT ID, DATE_END, CERT_NUMBER
FROM CERTS
WHERE UPPER(TRIM(FIO)) = UPPER(TRIM(@fio))
ORDER BY DATE_END DESC
ROWS 1";
                                    cmd.Parameters.AddWithValue("@fio", fioKey);

                                    using (var r = cmd.ExecuteReader())
                                    {
                                        if (r.Read())
                                        {
                                            existingId = r.GetInt32(0);
                                            existingDateEnd = r.IsDBNull(1) ? DateTime.MinValue : r.GetDateTime(1);
                                            existingCertNumber = r.IsDBNull(2) ? "" : r.GetString(2);
                                        }
                                    }
                                }
                            }

                            // Нормализуем номер для сравнения
                            string existingCertNorm = NormalizeCertNumber(existingCertNumber);
                            string newCertNorm = NormalizeCertNumber(entry.CertNumber);

                            if (existingId > 0)
                            {
                                bool sameCert = existingCertNorm == newCertNorm;
                                bool isNewer = entry.DateEnd > existingDateEnd;

                                if (sameCert || isNewer)
                                {
                                    using (var upd = conn.CreateCommand())
                                    {
                                        upd.Transaction = tx;
                                        upd.CommandText = @"
UPDATE CERTS SET
    FIO          = @fioOrig,
    CERT_NUMBER  = @certOrig,
    DATE_START   = @dateStart,
    DATE_END     = @dateEnd,
    DAYS_LEFT    = 0,  -- ← Больше не сохраняем вычисленное значение
    FROM_ADDRESS = @fromAddress,
    FOLDER_PATH  = @folderPath,
    MESSAGE_DATE = @msgDate
WHERE ID = @id";

                                        upd.Parameters.AddWithValue("@fioOrig", (object)entry.Fio ?? "");
                                        upd.Parameters.AddWithValue("@certOrig", (object)entry.CertNumber ?? "");
                                        upd.Parameters.AddWithValue("@dateStart",
                                            entry.DateStart == DateTime.MinValue ? (object)DBNull.Value : entry.DateStart);
                                        upd.Parameters.AddWithValue("@dateEnd",
                                            entry.DateEnd == DateTime.MinValue ? (object)DBNull.Value : entry.DateEnd);
                                        upd.Parameters.AddWithValue("@fromAddress", (object)entry.FromAddress ?? "");
                                        upd.Parameters.AddWithValue("@folderPath", (object)entry.FolderPath ?? "");
                                        upd.Parameters.AddWithValue("@msgDate",
                                            entry.MessageDate == DateTime.MinValue ? (object)DBNull.Value : entry.MessageDate);
                                        upd.Parameters.AddWithValue("@id", existingId);

                                        upd.ExecuteNonQuery();
                                    }

                                    certId = existingId;
                                    wasUpdated = true;

                                    if (!sameCert && isNewer)
                                    {
                                        Log($"InsertOrUpdateAndGetId: ЗАМЕНА сертификата для FIO='{entry.Fio}': " +
                                            $"старый номер={existingCertNumber}, новый={entry.CertNumber}, " +
                                            $"старый DATE_END={existingDateEnd:dd.MM.yyyy HH:mm:ss}, " +
                                            $"новый DATE_END={entry.DateEnd:dd.MM.yyyy HH:mm:ss}");
                                    }
                                    else if (sameCert)
                                    {
                                        Log($"InsertOrUpdateAndGetId: обновлены данные существующего сертификата для FIO='{entry.Fio}', номер={entry.CertNumber}");
                                    }
                                }
                                else
                                {
                                    // Новый сертификат старее или не лучше существующего — не трогаем
                                    certId = existingId;
                                    Log($"InsertOrUpdateAndGetId: найден более старый сертификат для FIO='{entry.Fio}', " +
                                        $"номер={entry.CertNumber}, DATE_END={entry.DateEnd:dd.MM.yyyy HH:mm:ss}. " +
                                        $"В БД уже есть более новый до {existingDateEnd:dd.MM.yyyy HH:mm:ss}. Запись не изменена.");
                                }
                            }
                            else
                            {
                                // 3) Для этого ФИО в базе ничего нет — обычный INSERT
                                using (var ins = conn.CreateCommand())
                                {
                                    ins.Transaction = tx;
                                    ins.CommandText = @"
INSERT INTO CERTS (
    FIO, CERT_NUMBER, DATE_START, DATE_END, DAYS_LEFT,
    FROM_ADDRESS, FOLDER_PATH, MESSAGE_DATE
) VALUES (
    @fioOrig, @certOrig, @dateStart, @dateEnd, @daysLeft,
    @fromAddress, @folderPath, @msgDate
)
RETURNING ID";

                                    ins.Parameters.AddWithValue("@fioOrig", (object)entry.Fio ?? "");
                                    ins.Parameters.AddWithValue("@certOrig", (object)entry.CertNumber ?? "");
                                    ins.Parameters.AddWithValue("@dateStart",
                                        entry.DateStart == DateTime.MinValue ? (object)DBNull.Value : entry.DateStart);
                                    ins.Parameters.AddWithValue("@dateEnd",
                                        entry.DateEnd == DateTime.MinValue ? (object)DBNull.Value : entry.DateEnd);
                                    ins.Parameters.AddWithValue("@daysLeft", 0);  // ← Сохраняем 0
                                    ins.Parameters.AddWithValue("@fromAddress", (object)entry.FromAddress ?? "");
                                    ins.Parameters.AddWithValue("@folderPath", (object)entry.FolderPath ?? "");
                                    ins.Parameters.AddWithValue("@msgDate",
                                        entry.MessageDate == DateTime.MinValue ? (object)DBNull.Value : entry.MessageDate);

                                    var insertedIdObj = ins.ExecuteScalar();
                                    if (insertedIdObj != null && insertedIdObj != DBNull.Value)
                                    {
                                        certId = Convert.ToInt32(insertedIdObj);
                                        wasAdded = true;
                                        Log($"InsertOrUpdateAndGetId: добавлен новый сертификат ID={certId} для FIO='{entry.Fio}', номер={entry.CertNumber}");
                                    }
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
                return (false, false, -1);
            }

            return (wasUpdated, wasAdded, certId);
        }

        #endregion

        #region Load / update UI helpers

        private CertRecord MapReaderToCertRecord(FbDataReader rdr)
        {
            var id = rdr.GetInt32(0);
            return new CertRecord
            {
                Id = id,
                Fio = rdr.IsDBNull(1) ? "" : rdr.GetString(1),
                DateStart = rdr.IsDBNull(2) ? DateTime.MinValue : rdr.GetDateTime(2),
                DateEnd = rdr.IsDBNull(3) ? DateTime.MinValue : rdr.GetDateTime(3),
                // DaysLeft из БД больше не используем – он всегда 0, считаешь в UI
                CertNumber = rdr.IsDBNull(5) ? "" : rdr.GetString(5),
                FromAddress = rdr.IsDBNull(6) ? "" : rdr.GetString(6),
                IsDeleted = !rdr.IsDBNull(7) && rdr.GetInt16(7) == 1,
                Note = rdr.IsDBNull(8) ? "" : rdr.GetString(8),
                Building = rdr.IsDBNull(9) ? "" : rdr.GetString(9),
                FolderPath = rdr.IsDBNull(10) ? "" : rdr.GetString(10),
                ArchivePath = rdr.IsDBNull(11) ? "" : rdr.GetString(11),
                MessageDate = rdr.IsDBNull(12) ? DateTime.MinValue : rdr.GetDateTime(12),

                // новое поле из SELECT (HAS_ARCHIVE)
                HasArchive = !rdr.IsDBNull(13) && rdr.GetInt32(13) > 0
            };
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
                    if (!showDeleted) where.Add("c.IS_DELETED = 0");
                    if (!string.IsNullOrEmpty(buildingFilter) && buildingFilter != "Все")
                    {
                        where.Add("c.BUILDING = @building");
                        cmd.Parameters.AddWithValue("@building", buildingFilter);
                    }

                    // Показываем только не истёкшие
                    where.Add("c.DATE_END >= @currentDate");
                    cmd.Parameters.AddWithValue("@currentDate", DateTime.Now.Date);

                    string whereClause = where.Count > 0 ? "WHERE " + string.Join(" AND ", where) : "";

                    cmd.CommandText = $@"
SELECT 
    c.ID,
    c.FIO,
    c.DATE_START,
    c.DATE_END,
    c.DAYS_LEFT,
    c.CERT_NUMBER,
    c.FROM_ADDRESS,
    c.IS_DELETED,
    c.NOTE,
    c.BUILDING,
    c.FOLDER_PATH,
    c.ARCHIVE_PATH,
    c.MESSAGE_DATE,
    (SELECT COUNT(*) FROM CERT_ARCHIVES ca WHERE ca.CERT_ID = c.ID) AS HAS_ARCHIVE
FROM CERTS c
{whereClause}
ORDER BY c.DATE_END ASC";

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


            Log($"LoadAll: {list.Count} записей, showDeleted={showDeleted}, filter={buildingFilter}, время={sw.ElapsedMilliseconds}мс");
            // Log($"LoadAll: загружено {list.Count} записей (showDeleted={showDeleted}, filter={buildingFilter})");
            return list;
        }

        public void UpdateNote(int id, string note)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET NOTE = @note WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@note", note ?? "");
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
            Log($"UpdateNote: ID={id}");
        }

        public void UpdateBuilding(int id, string building)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET BUILDING = @b WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@b", building ?? "");
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
            Log($"UpdateBuilding: ID={id}, building={building}");
        }

        public void MarkAsDeleted(int id, bool isDeleted)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET IS_DELETED = @d WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@d", isDeleted ? 1 : 0);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
            Log($"MarkAsDeleted: ID={id}, isDeleted={isDeleted}");
        }

        public void UpdateArchivePath(int id, string archivePath)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET ARCHIVE_PATH = @p WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@p", archivePath ?? "");
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                }
            }
            Log($"UpdateArchivePath: ID={id}, path={archivePath}");
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
            if (string.IsNullOrEmpty(mailUid)) return;
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
INSERT INTO PROCESSED_MAILS (FOLDER_PATH, MAIL_UID, KIND, PROCESSED_DATE)
VALUES (@folder, @uid, @kind, @dt)";
                        cmd.Parameters.AddWithValue("@folder", folderPath);
                        cmd.Parameters.AddWithValue("@uid", mailUid);
                        cmd.Parameters.AddWithValue("@kind", kind);
                        cmd.Parameters.AddWithValue("@dt", DateTime.Now);

                        try
                        {
                            cmd.ExecuteNonQuery();
                            Log($"MarkMailProcessed: folder='{folderPath}', uid='{mailUid}', kind='{kind}'");
                        }
                        catch (FbException fbEx)
                        {
                            // Если это нарушение уникального индекса — значит уже помечали, просто игнорируем.
                            Log($"MarkMailProcessed: уже существует запись для UID='{mailUid}', kind='{kind}' ({fbEx.Message})");
                        }
                        catch (Exception ex)
                        {
                            Log($"MarkMailProcessed error: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception exOuter)
            {
                Log($"MarkMailProcessed outer error: {exOuter.Message}");
            }
        }

        #endregion


        #region Revocation helpers

        /// <summary>
        /// Находит сертификат по номеру и помечает как аннулированный, добавляя в NOTE дату отзыва (если указана).
        /// Возвращает true если найден и обновлён.
        /// </summary>
        public bool FindAndMarkAsRevokedByCertNumber(string certNumber, string fio, string folderPath, string revokeDateShort)
        {
            Log($"FindAndMarkAsRevokedByCertNumber: cert={certNumber}, fio={fio}, date={revokeDateShort}");
            try
            {
                var certNorm = NormalizeCertNumber(certNumber);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT ID, NOTE FROM CERTS WHERE UPPER(TRIM(CERT_NUMBER)) = @cert";
                        cmd.Parameters.AddWithValue("@cert", certNorm);

                        using (var rdr = cmd.ExecuteReader())
                        {
                            if (!rdr.Read())
                            {
                                Log($"FindAndMarkAsRevokedByCertNumber: сертификат не найден {certNumber}");
                                return false;
                            }

                            int id = rdr.GetInt32(0);
                            string oldNote = rdr.IsDBNull(1) ? "" : rdr.GetString(1);
                            rdr.Close();

                            string newNote = !string.IsNullOrEmpty(revokeDateShort) ? $"аннулирован ({revokeDateShort})" : "аннулирован";
                            string noteToSet = string.IsNullOrEmpty(oldNote) ? newNote : $"{oldNote}; {newNote}";

                            using (var upd = conn.CreateCommand())
                            {
                                upd.CommandText = "UPDATE CERTS SET IS_DELETED = 1, NOTE = @note, FOLDER_PATH = @folder WHERE ID = @id";
                                upd.Parameters.AddWithValue("@note", noteToSet);
                                upd.Parameters.AddWithValue("@folder", folderPath ?? "");
                                upd.Parameters.AddWithValue("@id", id);
                                upd.ExecuteNonQuery();
                            }

                            Log($"FindAndMarkAsRevokedByCertNumber: аннулирован ID={id}, cert={certNumber}");
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"FindAndMarkAsRevokedByCertNumber error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Находит по FIO+certNumber и помечает аннулированным.
        /// </summary>
        public bool FindAndMarkAsRevoked(string fio, string certNumber, string folderPath)
        {
            try
            {
                var fioNorm = NormalizeFio(fio);
                var certNorm = NormalizeCertNumber(certNumber);

                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();
                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "SELECT ID FROM CERTS WHERE UPPER(TRIM(FIO)) = @fio AND UPPER(TRIM(CERT_NUMBER)) = @cert";
                        cmd.Parameters.AddWithValue("@fio", fioNorm);
                        cmd.Parameters.AddWithValue("@cert", certNorm);

                        var r = cmd.ExecuteScalar();
                        if (r == null) return false;

                        int id = Convert.ToInt32(r);
                        using (var upd = conn.CreateCommand())
                        {
                            upd.CommandText = "UPDATE CERTS SET IS_DELETED = 1, NOTE = CASE WHEN NOTE IS NULL OR NOTE = '' THEN 'аннулирован' ELSE NOTE || '; аннулирован' END, FOLDER_PATH = @f WHERE ID = @id";
                            upd.Parameters.AddWithValue("@f", folderPath ?? "");
                            upd.Parameters.AddWithValue("@id", id);
                            upd.ExecuteNonQuery();
                        }
                        Log($"FindAndMarkAsRevoked: аннулирован ID={id}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"FindAndMarkAsRevoked error: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Logging

        private void Log(string message)
        {
            try
            {
                string logFile = ImapCertWatcher.Utils.LogSession.SessionLogFile;
                string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [DbHelper] {message}{Environment.NewLine}";
                File.AppendAllText(logFile, entry, System.Text.Encoding.UTF8);

                _addToMiniLog?.Invoke($"[DbHelper] {message}");
                System.Diagnostics.Debug.WriteLine($"[DbHelper] {message}");
            }
            catch
            {
                // не фейлим приложение из-за логирования
            }
        }

        #endregion
    }
}

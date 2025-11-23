using FirebirdSql.Data.FirebirdClient;
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using System;
using System.Collections.Generic;
using System.IO;

namespace ImapCertWatcher.Data
{
    public class DbHelper
    {
        private readonly AppSettings _settings;
        private readonly string _connectionString;
        private readonly string _logDirectory;
        private readonly Action<string> _addToMiniLog;

        public DbHelper(AppSettings settings, Action<string> addToMiniLog = null)
        {
            _settings = settings;
            _connectionString = BuildConnectionString();
            _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
            _addToMiniLog = addToMiniLog;

            if (!Directory.Exists(_logDirectory))
                Directory.CreateDirectory(_logDirectory);

            EnsureDatabaseAndTable();
            Log("Инициализация DbHelper завершена");
        }

        private string BuildConnectionString()
        {
            var csb = new FbConnectionStringBuilder
            {
                Database = _settings.FirebirdDbPath,
                UserID = _settings.FbUser,
                Password = _settings.FbPassword,
                DataSource = _settings.FbServer,
                Port = 3050,
                Charset = "UTF8",
                Dialect = _settings.FbDialect,
                ServerType = FbServerType.Default
            };

            Log($"Строка подключения к БД: Server={_settings.FbServer}, Database={_settings.FirebirdDbPath}");
            return csb.ToString();
        }

        public void EnsureDatabaseAndTable()
        {
            if (_settings.FbServer == "127.0.0.1" || _settings.FbServer == "localhost")
            {
                string dir = Path.GetDirectoryName(_settings.FirebirdDbPath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                if (!File.Exists(_settings.FirebirdDbPath))
                {
                    var csb = new FbConnectionStringBuilder(_connectionString);
                    FbConnection.CreateDatabase(csb.ToString(), pageSize: 16384, forcedWrites: true, overwrite: false);
                    Log("Локальная база данных создана");
                }
            }

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                bool tableExists = CheckTableExists(conn);

                if (!tableExists)
                {
                    CreateTable(conn);
                }
                else
                {
                    AddMissingColumns(conn);
                }
            }
        }

        private bool CheckTableExists(FbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    @"SELECT COUNT(*) 
                      FROM RDB$RELATIONS 
                      WHERE RDB$RELATION_NAME = 'CERTS' 
                      AND RDB$SYSTEM_FLAG = 0";

                var result = cmd.ExecuteScalar();
                bool exists = Convert.ToInt32(result) > 0;
                Log($"Проверка существования таблицы CERTS: {exists}");
                return exists;
            }
        }

        private void CreateTable(FbConnection conn)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    @"CREATE TABLE CERTS (
                        ID INTEGER NOT NULL PRIMARY KEY,
                        FIO VARCHAR(200),
                        DATE_START TIMESTAMP,
                        DATE_END TIMESTAMP,
                        DAYS_LEFT INTEGER,
                        CERT_NUMBER VARCHAR(100),
                        FROM_ADDRESS VARCHAR(200),
                        IS_DELETED SMALLINT DEFAULT 0,
                        NOTE VARCHAR(500),
                        BUILDING VARCHAR(100),
                        FOLDER_PATH VARCHAR(300),
                        ARCHIVE_PATH VARCHAR(500),
                        MESSAGE_DATE TIMESTAMP
                    );";

                cmd.ExecuteNonQuery();

                cmd.CommandText = "CREATE SEQUENCE SEQ_CERTS_ID;";
                cmd.ExecuteNonQuery();

                cmd.CommandText =
                    @"CREATE TRIGGER CERTS_BI FOR CERTS
                      ACTIVE BEFORE INSERT POSITION 0
                      AS
                      BEGIN
                        IF (NEW.ID IS NULL) THEN
                          NEW.ID = NEXT VALUE FOR SEQ_CERTS_ID;
                      END";

                cmd.ExecuteNonQuery();

                Log("Таблица CERTS создана успешно");
            }
        }

        private void AddMissingColumns(FbConnection conn)
        {
            var columnsToAdd = new Dictionary<string, string>
            {
                { "CERT_NUMBER", "VARCHAR(100)" },
                { "FROM_ADDRESS", "VARCHAR(200)" },
                { "IS_DELETED", "SMALLINT DEFAULT 0" },
                { "NOTE", "VARCHAR(500)" },
                { "BUILDING", "VARCHAR(100)" },
                { "FOLDER_PATH", "VARCHAR(300)" },
                { "ARCHIVE_PATH", "VARCHAR(500)" },
                { "MESSAGE_DATE", "TIMESTAMP" }
            };

            foreach (var column in columnsToAdd)
            {
                if (!ColumnExists(conn, column.Key))
                {
                    try
                    {
                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = $@"ALTER TABLE CERTS ADD {column.Key} {column.Value}";
                            cmd.ExecuteNonQuery();
                            Log($"Добавлен столбец: {column.Key}");
                        }
                    }
                    catch (FbException ex)
                    {
                        Log($"Ошибка при добавлении столбца {column.Key}: {ex.Message}");
                    }
                }
                else
                {
                    if (column.Key == "BUILDING")
                    {
                        TryResizeBuildingColumn(conn);
                    }
                }
            }
        }

        public bool FindAndMarkAsRevokedByCertNumber(string certNumber, string fio, string folderPath)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        // Ищем только по номеру сертификата — на 100% надёжно
                        cmd.CommandText = @"SELECT ID FROM CERTS WHERE CERT_NUMBER = @certNumber AND IS_DELETED = 0";
                        cmd.Parameters.AddWithValue("@certNumber", certNumber);

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                int id = reader.GetInt32(0);
                                reader.Close();

                                cmd.Parameters.Clear();
                                cmd.CommandText = @"
                            UPDATE CERTS 
                            SET IS_DELETED = 1,
                                NOTE = CASE 
                                    WHEN NOTE IS NULL OR NOTE = '' THEN 'аннулирован'
                                    ELSE NOTE || '; аннулирован'
                                END,
                                FOLDER_PATH = @folderPath
                            WHERE ID = @id";

                                cmd.Parameters.AddWithValue("@folderPath", folderPath ?? "");
                                cmd.Parameters.AddWithValue("@id", id);

                                cmd.ExecuteNonQuery();
                                Log($"Сертификат аннулирован (по номеру): {certNumber}");
                                return true;
                            }
                        }

                        Log($"Сертификат не найден (по номеру): {certNumber}");
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Log($"Ошибка FindAndMarkAsRevokedByCertNumber: {ex.Message}");
                return false;
            }
        }


        public bool FindAndMarkAsRevoked(string fio, string certNumber, string folderPath)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"SELECT ID FROM CERTS 
                                    WHERE FIO = @fio AND CERT_NUMBER = @certNumber
                                    AND IS_DELETED = 0";
                        cmd.Parameters.AddWithValue("@fio", fio);
                        cmd.Parameters.AddWithValue("@certNumber", certNumber ?? "");

                        using (var reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                int id = reader.GetInt32(0);
                                reader.Close();

                                // Обновляем запись - помечаем как удаленную и добавляем примечание
                                cmd.Parameters.Clear();
                                cmd.CommandText = @"UPDATE CERTS 
                                            SET IS_DELETED = 1, 
                                                NOTE = CASE 
                                                    WHEN NOTE IS NULL OR NOTE = '' THEN 'аннулирован' 
                                                    ELSE NOTE || '; аннулирован' 
                                                END,
                                                FOLDER_PATH = @folderPath
                                            WHERE ID = @id";

                                cmd.Parameters.AddWithValue("@folderPath", folderPath ?? "");
                                cmd.Parameters.AddWithValue("@id", id);

                                int rowsAffected = cmd.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    Log($"Сертификат помечен как аннулированный: ID={id}, {fio}, {certNumber}");
                                    return true;
                                }
                            }
                            else
                            {
                                Log($"Сертификат не найден для аннулирования: {fio}, {certNumber}");
                            }
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Log($"Ошибка при пометке сертификата как аннулированного: {ex.Message}");
                return false;
            }
        }

        public bool CheckDuplicate(string fio, string certNumber)
        {
            try
            {
                using (var conn = new FbConnection(_connectionString))
                {
                    conn.Open();

                    using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"SELECT COUNT(*) FROM CERTS 
                                    WHERE FIO = @fio AND CERT_NUMBER = @certNumber";
                        cmd.Parameters.AddWithValue("@fio", fio);
                        cmd.Parameters.AddWithValue("@certNumber", certNumber ?? "");

                        var result = cmd.ExecuteScalar();
                        bool exists = Convert.ToInt32(result) > 0;

                        if (exists)
                        {
                            Log($"Найден дубликат: {fio}, сертификат: {certNumber}");
                        }
                        else
                        {
                            Log($"Дубликат не найден: {fio}, {certNumber}");
                        }

                        return exists;
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Ошибка проверки дубликата: {ex.Message}");
                return false; // При ошибке лучше пропустить проверку
            }
        }

        private void TryResizeBuildingColumn(FbConnection conn)
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"ALTER TABLE CERTS ALTER COLUMN BUILDING TYPE VARCHAR(100)";
                    cmd.ExecuteNonQuery();
                    
                }
            }
            catch (FbException ex)
            {
                Log($"Не удалось изменить размер столбца BUILDING: {ex.Message}");
            }
        }

        private bool ColumnExists(FbConnection conn, string columnName)
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    @"SELECT COUNT(*) 
                      FROM RDB$RELATION_FIELDS 
                      WHERE RDB$RELATION_NAME = 'CERTS' 
                      AND RDB$FIELD_NAME = @columnName";
                cmd.Parameters.AddWithValue("@columnName", columnName);

                var result = cmd.ExecuteScalar();
                return Convert.ToInt32(result) > 0;
            }
        }

        public (bool wasUpdated, bool wasAdded) InsertOrUpdate(CertEntry entry)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"SELECT ID, MESSAGE_DATE FROM CERTS WHERE FIO = @fio AND CERT_NUMBER = @certNumber";
                    cmd.Parameters.AddWithValue("@fio", entry.Fio);
                    cmd.Parameters.AddWithValue("@certNumber", entry.CertNumber ?? "");

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int existingId = reader.GetInt32(0);

                            DateTime existingDate = DateTime.MinValue;
                            if (!reader.IsDBNull(1))
                            {
                                existingDate = reader.GetDateTime(1);
                            }

                            reader.Close();

                            // Если письмо новее существующего
                            if (entry.MessageDate > existingDate)
                            {
                                cmd.Parameters.Clear();
                                cmd.CommandText =
                                    @"UPDATE CERTS SET 
                                DATE_START = @ds,
                                DATE_END = @de,
                                DAYS_LEFT = @dleft,
                                FROM_ADDRESS = @fromAddress,
                                FOLDER_PATH = @folderPath,
                                ARCHIVE_PATH = @archivePath,
                                MESSAGE_DATE = @messageDate
                              WHERE ID = @id";

                                cmd.Parameters.AddWithValue("@ds", entry.DateStart);
                                cmd.Parameters.AddWithValue("@de", entry.DateEnd);
                                cmd.Parameters.AddWithValue("@dleft", entry.DaysLeft);
                                cmd.Parameters.AddWithValue("@fromAddress", entry.FromAddress ?? "");
                                cmd.Parameters.AddWithValue("@folderPath", entry.FolderPath ?? "");
                                cmd.Parameters.AddWithValue("@archivePath", entry.ArchivePath ?? "");
                                cmd.Parameters.AddWithValue("@messageDate", entry.MessageDate);
                                cmd.Parameters.AddWithValue("@id", existingId);

                                cmd.ExecuteNonQuery();
                                Log($"Обновлена запись (более новое письмо): {entry.Fio}, сертификат: {entry.CertNumber}");
                                return (true, false); // Было обновление
                            }
                            else
                            {
                                Log($"Пропущена запись (уже есть более новое письмо): {entry.Fio}, сертификат: {entry.CertNumber}");
                                return (false, false); // Не было изменений
                            }
                        }
                        else
                        {
                            reader.Close();

                            cmd.Parameters.Clear();
                            cmd.CommandText =
                                @"INSERT INTO CERTS (FIO, DATE_START, DATE_END, DAYS_LEFT, CERT_NUMBER, FROM_ADDRESS, IS_DELETED, NOTE, BUILDING, FOLDER_PATH, ARCHIVE_PATH, MESSAGE_DATE)
                  VALUES (@fio, @ds, @de, @dleft, @certNumber, @fromAddress, @isDeleted, @note, @building, @folderPath, @archivePath, @messageDate)";

                            cmd.Parameters.AddWithValue("@fio", entry.Fio);
                            cmd.Parameters.AddWithValue("@ds", entry.DateStart);
                            cmd.Parameters.AddWithValue("@de", entry.DateEnd);
                            cmd.Parameters.AddWithValue("@dleft", entry.DaysLeft);
                            cmd.Parameters.AddWithValue("@certNumber", entry.CertNumber ?? "");
                            cmd.Parameters.AddWithValue("@fromAddress", entry.FromAddress ?? "");
                            cmd.Parameters.AddWithValue("@isDeleted", 0);
                            cmd.Parameters.AddWithValue("@note", "");
                            cmd.Parameters.AddWithValue("@building", "");
                            cmd.Parameters.AddWithValue("@folderPath", entry.FolderPath ?? "");
                            cmd.Parameters.AddWithValue("@archivePath", entry.ArchivePath ?? "");
                            cmd.Parameters.AddWithValue("@messageDate", entry.MessageDate);

                            cmd.ExecuteNonQuery();
                            Log($"Добавлена новая запись: {entry.Fio}, сертификат: {entry.CertNumber}");
                            return (false, true); // Было добавление
                        }
                    }
                }
            }
        }

        public List<CertRecord> LoadAll(bool showDeleted = false, string buildingFilter = "Все")
        {
            var list = new List<CertRecord>();
            Log($"Загрузка данных из БД: showDeleted={showDeleted}, buildingFilter={buildingFilter}");

            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    var whereConditions = new List<string>();
                    var parameters = new Dictionary<string, object>();

                    if (!showDeleted)
                    {
                        whereConditions.Add("IS_DELETED = 0");
                    }

                    if (buildingFilter != "Все")
                    {
                        whereConditions.Add("BUILDING = @building");
                        parameters.Add("@building", buildingFilter);
                    }

                    whereConditions.Add("DATE_END >= @currentDate");
                    parameters.Add("@currentDate", DateTime.Now.Date);

                    string whereClause = whereConditions.Count > 0 ? "WHERE " + string.Join(" AND ", whereConditions) : "";

                    cmd.CommandText =
                        $@"SELECT ID, FIO, DATE_START, DATE_END, DAYS_LEFT, CERT_NUMBER, FROM_ADDRESS, IS_DELETED, NOTE, BUILDING, FOLDER_PATH, ARCHIVE_PATH, MESSAGE_DATE
                  FROM CERTS
                  {whereClause}
                  ORDER BY DATE_END ASC";

                    foreach (var param in parameters)
                    {
                        cmd.Parameters.AddWithValue(param.Key, param.Value);
                    }

                    using (var reader = cmd.ExecuteReader())
                    {
                        int count = 0;
                        while (reader.Read())
                        {
                            var building = reader.IsDBNull(9) ? "" : reader.GetString(9);
                            var record = new CertRecord
                            {
                                Id = reader.GetInt32(0),
                                Fio = reader.IsDBNull(1) ? "" : reader.GetString(1),
                                DateStart = reader.GetDateTime(2),
                                DateEnd = reader.GetDateTime(3),
                                DaysLeft = reader.GetInt32(4),
                                CertNumber = reader.IsDBNull(5) ? "" : reader.GetString(5),
                                FromAddress = reader.IsDBNull(6) ? "" : reader.GetString(6),
                                IsDeleted = !reader.IsDBNull(7) && reader.GetInt16(7) == 1,
                                Note = reader.IsDBNull(8) ? "" : reader.GetString(8),
                                Building = building,
                                FolderPath = reader.IsDBNull(10) ? "" : reader.GetString(10),
                                ArchivePath = reader.IsDBNull(11) ? "" : reader.GetString(11),
                                MessageDate = reader.IsDBNull(12) ? DateTime.MinValue : reader.GetDateTime(12)
                            };

                            list.Add(record);
                            count++;

                            
                        }
                        Log($"Загружено записей из БД: {count}");
                    }
                }
            }

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
                    Log($"Обновлено примечание для записи ID={id}");
                }
            }
        }

        public void UpdateBuilding(int id, string building)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET BUILDING = @building WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@building", building ?? "");
                    cmd.Parameters.AddWithValue("@id", id);

                    int rowsAffected = cmd.ExecuteNonQuery();
                    Log($"Обновлено здание для записи ID={id}: '{building}' (затронуто строк: {rowsAffected})");

                    if (rowsAffected == 0)
                    {
                        Log($"ВНИМАНИЕ: Не найдена запись с ID={id} для обновления здания");
                    }
                }
            }
        }

        public void MarkAsDeleted(int id, bool isDeleted)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET IS_DELETED = @isDeleted WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@isDeleted", isDeleted ? 1 : 0);
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                    Log($"Изменен статус удаления для записи ID={id}: {isDeleted}");
                }
            }
        }

        public void UpdateArchivePath(int id, string archivePath)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "UPDATE CERTS SET ARCHIVE_PATH = @archivePath WHERE ID = @id";
                    cmd.Parameters.AddWithValue("@archivePath", archivePath ?? "");
                    cmd.Parameters.AddWithValue("@id", id);
                    cmd.ExecuteNonQuery();
                    Log($"Обновлен путь к архиву для записи ID={id}: {archivePath}");
                }
            }
        }

        public FbConnection GetConnection() => new FbConnection(_connectionString);

        private void Log(string message)
        {
            try
            {
                // Один файл на день вместо файла на каждую секунду
                string logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG", DateTime.Now.ToString("yyyy-MM-dd"));
                if (!Directory.Exists(logDirectory))
                    Directory.CreateDirectory(logDirectory);

                // Файл с именем дня
                string logFile = Path.Combine(logDirectory, $"log_{DateTime.Now:yyyy-MM-dd}.log");
                string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [ImapWatcher] {message}{Environment.NewLine}";

                File.AppendAllText(logFile, logEntry, System.Text.Encoding.UTF8);

                // Также пишем в мини-лог через переданный делегат
                _addToMiniLog?.Invoke($"[ImapWatcher] {message}");

                System.Diagnostics.Debug.WriteLine($"[ImapWatcher] {message}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка записи в лог: {ex.Message}");
            }
        }
    }
}
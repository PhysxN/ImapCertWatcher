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

        public DbHelper(AppSettings settings)
        {
            _settings = settings;
            _connectionString = BuildConnectionString();
            EnsureDatabaseAndTable();
        }

        private string BuildConnectionString()
        {
            var csb = new FbConnectionStringBuilder
            {
                Database = _settings.FirebirdDbPath,
                UserID = _settings.FbUser,
                Password = _settings.FbPassword,
                DataSource = "localhost",
                Charset = "UTF8",
                Dialect = _settings.FbDialect,
                ServerType = FbServerType.Default
            };

            return csb.ToString();
        }

        public void EnsureDatabaseAndTable()
        {
            string dir = Path.GetDirectoryName(_settings.FirebirdDbPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (!File.Exists(_settings.FirebirdDbPath))
            {
                var csb = new FbConnectionStringBuilder(_connectionString);
                FbConnection.CreateDatabase(csb.ToString(), pageSize: 16384, forcedWrites: true, overwrite: false);
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
                return Convert.ToInt32(result) > 0;
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
                        BUILDING VARCHAR(50),
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

                System.Diagnostics.Debug.WriteLine("Таблица CERTS создана успешно");
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
                { "BUILDING", "VARCHAR(50)" },
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
                            System.Diagnostics.Debug.WriteLine($"Добавлен столбец: {column.Key}");
                        }
                    }
                    catch (FbException ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Ошибка при добавлении столбца {column.Key}: {ex.Message}");
                    }
                }
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

        public void InsertOrUpdate(CertEntry entry)
        {
            using (var conn = new FbConnection(_connectionString))
            {
                conn.Open();

                using (var cmd = conn.CreateCommand())
                {
                    // Ищем существующую запись с таким же ФИО
                    cmd.CommandText = @"SELECT ID, MESSAGE_DATE FROM CERTS WHERE FIO = @fio AND CERT_NUMBER = @certNumber";
                    cmd.Parameters.AddWithValue("@fio", entry.Fio);
                    cmd.Parameters.AddWithValue("@certNumber", entry.CertNumber ?? "");

                    using (var reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            // Запись существует - проверяем дату
                            int existingId = reader.GetInt32(0);
                            DateTime existingDate = reader.GetDateTime(1);

                            // Если новое письмо более свежее - обновляем
                            if (entry.MessageDate > existingDate)
                            {
                                reader.Close();

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
                                System.Diagnostics.Debug.WriteLine($"Обновлена запись (более новое письмо): {entry.Fio}, сертификат: {entry.CertNumber}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"Пропущена запись (уже есть более новое письмо): {entry.Fio}, сертификат: {entry.CertNumber}");
                            }
                        }
                        else
                        {
                            reader.Close();

                            // Новая запись
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
                            System.Diagnostics.Debug.WriteLine($"Добавлена новая запись: {entry.Fio}, сертификат: {entry.CertNumber}");
                        }
                    }
                }
            }
        }

        public List<CertRecord> LoadAll(bool showDeleted = false, string buildingFilter = "Все")
        {
            var list = new List<CertRecord>();

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
                        while (reader.Read())
                        {
                            list.Add(new CertRecord
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
                                Building = reader.IsDBNull(9) ? "" : reader.GetString(9),
                                FolderPath = reader.IsDBNull(10) ? "" : reader.GetString(10),
                                ArchivePath = reader.IsDBNull(11) ? "" : reader.GetString(11),
                                MessageDate = reader.IsDBNull(12) ? DateTime.MinValue : reader.GetDateTime(12)
                            });
                        }
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
                    cmd.ExecuteNonQuery();
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
                }
            }
        }

        public FbConnection GetConnection() => new FbConnection(_connectionString);
    }
}
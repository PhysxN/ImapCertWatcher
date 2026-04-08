using ImapCertWatcher.Models;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private void BtnLoadCer_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(
                "Ручная загрузка сертификата временно отключена.",
                "Временно недоступно",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private async void AddArchiveMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (!(dgCerts.SelectedItem is CertRecord record))
            {
                MessageBox.Show("Не выбрана запись");
                return;
            }

            var dlg = new OpenFileDialog
            {
                Filter = "ZIP (*.zip)|*.zip"
            };

            if (dlg.ShowDialog() != true)
                return;

            try
            {
                var fileInfo = new FileInfo(dlg.FileName);

                if (!fileInfo.Exists)
                {
                    MessageBox.Show(
                        "Файл не найден. Возможно, он был удален или перемещен.",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (fileInfo.Length <= 0)
                {
                    MessageBox.Show(
                        "Файл пустой.",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                // такой же предел, как и на серверной обработке ZIP
                if (fileInfo.Length > 50_000_000)
                {
                    MessageBox.Show(
                        "Файл слишком большой. Максимальный размер — 50 МБ.",
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                byte[] bytes = File.ReadAllBytes(dlg.FileName);
                string base64 = Convert.ToBase64String(bytes);
                string fileName = Path.GetFileName(dlg.FileName);

                var response = await _api.AddArchive(record.Id, fileName, base64);

                if (string.IsNullOrWhiteSpace(response))
                {
                    MessageBox.Show("Нет ответа от сервера");
                    return;
                }

                if (!response.StartsWith("OK"))
                {
                    MessageBox.Show("Ошибка добавления архива:\n" + response);
                    return;
                }

                record.HasArchive = true;

                AddToMiniLog($"Архив {fileName} добавлен к записи {record.Fio}");
                statusText.Text = "Архив успешно добавлен";
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(
                    "Файл не найден. Возможно, он был удален или перемещен.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(
                    "Нет прав доступа к файлу.",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (IOException ex)
            {
                MessageBox.Show(
                    $"Ошибка ввода-вывода при чтении файла:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Неожиданная ошибка при чтении файла:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private async void BtnOpenArchive_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (Mouse.OverrideCursor == Cursors.Wait)
                return;

            if (!(dgCerts.SelectedItem is CertRecord record))
            {
                MessageBox.Show("Не выбрана запись");
                return;
            }

            try
            {
                int certId = record.Id;
                string fio = record.Fio ?? "";
                string certFolder = GetCertFolderPath(record);

                bool hasFiles = false;

                try
                {
                    hasFiles = Directory.Exists(certFolder) &&
                               Directory.EnumerateFiles(certFolder, "*", SearchOption.AllDirectories).Any();
                }
                catch
                {
                }

                if (hasFiles)
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", certFolder)
                    {
                        UseShellExecute = true
                    });

                    statusText.Text = "Архив уже был загружен ранее";
                    return;
                }

                Mouse.OverrideCursor = Cursors.Wait;
                statusText.Text = "Загрузка архива...";

                var response = await _api.SendCommand($"GET_ARCHIVE|{certId}");

                if (string.IsNullOrWhiteSpace(response) || !response.StartsWith("ARCHIVE|"))
                {
                    MessageBox.Show("Архив не найден");
                    statusText.Text = "Архив не найден";
                    return;
                }

                var payload = response.Substring("ARCHIVE|".Length);
                var parts = payload.Split(new[] { '|' }, 2);

                if (parts.Length < 2)
                {
                    MessageBox.Show("Некорректный формат архива");
                    statusText.Text = "Ошибка формата архива";
                    return;
                }

                string fileName = Path.GetFileName(parts[0]);

                if (string.IsNullOrWhiteSpace(fileName))
                {
                    MessageBox.Show("Некорректное имя файла в архиве");
                    statusText.Text = "Ошибка имени файла";
                    return;
                }

                byte[] fileData;
                try
                {
                    fileData = await Task.Run(() => Convert.FromBase64String(parts[1]));
                }
                catch
                {
                    MessageBox.Show("Повреждённые данные архива");
                    statusText.Text = "Ошибка чтения архива";
                    return;
                }

                string tempZip = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".zip");

                try
                {
                    await Task.Run(() => File.WriteAllBytes(tempZip, fileData));

                    statusText.Text = "Распаковка архива...";

                    if (fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                    {
                        string finalFolder = certFolder;

                        await Task.Run(() =>
                        {
                            if (Directory.Exists(finalFolder))
                            {
                                try
                                {
                                    Directory.Delete(finalFolder, true);
                                }
                                catch
                                {
                                    finalFolder = finalFolder + "_" + DateTime.Now.Ticks;
                                }
                            }

                            Directory.CreateDirectory(finalFolder);
                            ZipFile.ExtractToDirectory(tempZip, finalFolder);
                        });

                        certFolder = finalFolder;
                    }
                    else
                    {
                        if (Directory.Exists(certFolder))
                        {
                            try
                            {
                                Directory.Delete(certFolder, true);
                            }
                            catch
                            {
                                certFolder = certFolder + "_" + DateTime.Now.Ticks;
                            }
                        }

                        Directory.CreateDirectory(certFolder);

                        string filePath = Path.Combine(certFolder, fileName);
                        await Task.Run(() => File.WriteAllBytes(filePath, fileData));
                    }

                    Process.Start(new ProcessStartInfo("explorer.exe", certFolder)
                    {
                        UseShellExecute = true
                    });

                    AddToMiniLog($"Архив для {fio} загружен и распакован");
                    statusText.Text = "Архив успешно загружен";
                }
                finally
                {
                    if (File.Exists(tempZip))
                    {
                        try { File.Delete(tempZip); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка открытия сертификата:\n" + ex.Message);
                statusText.Text = "Ошибка открытия архива";
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private string GetCertsRoot()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Certs");
        }

        private async Task CleanupCertsFolderAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    string root = GetCertsRoot();

                    Directory.CreateDirectory(root);

                    foreach (var dir in Directory.GetDirectories(root))
                    {
                        try
                        {
                            var info = new DirectoryInfo(dir);
                            var age = DateTime.Now - info.LastWriteTime;

                            if (age > TimeSpan.FromHours(48))
                            {
                                Directory.Delete(dir, true);
                            }
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                }
            });
        }

        private string GetCertFolderPath(CertRecord record)
        {
            if (record == null)
                throw new ArgumentNullException(nameof(record));

            string safeName = MakeSafeFolderName(record.Fio);
            string folderName = $"{record.Id}_{safeName}";

            return Path.Combine(GetCertsRoot(), folderName);
        }


        private string MakeSafeFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "NO_NAME";

            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            name = name.Trim();

            if (string.IsNullOrWhiteSpace(name))
                return "NO_NAME";

            return name;
        }
    }
}
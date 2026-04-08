using ImapCertWatcher.Models;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace ImapCertWatcher
{
    public partial class MainWindow
    {
        private async void BtnAddToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            var sn = Prompt.ShowDialog(
                "Введите серийный номер токена",
                "Добавление токена");

            if (string.IsNullOrWhiteSpace(sn))
                return;

            sn = sn.Trim().ToUpperInvariant();

            var response = await _api.AddToken(sn);

            if (string.IsNullOrWhiteSpace(response))
            {
                MessageBox.Show("Нет ответа от сервера");
                return;
            }

            response = response.Trim();

            if (response.StartsWith("ERROR|ADD_TOKEN|TOKEN_ALREADY_EXISTS"))
            {
                MessageBox.Show(
                    "Токен с таким серийным номером уже существует.",
                    "Дубликат",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (response.StartsWith("ERROR|ADD_TOKEN|"))
            {
                MessageBox.Show(
                    "Ошибка добавления токена:\n" + response.Substring("ERROR|ADD_TOKEN|".Length),
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (!response.StartsWith("OK"))
            {
                MessageBox.Show(
                    "Ошибка добавления токена:\n" + response,
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            await LoadFromServer();
            AddToMiniLog($"Токен {sn} добавлен");
        }
        private async void BtnUnassignToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            if (!TryGetToken(sender, out var token))
                return;

            if (token.OwnerCertId == null)
            {
                MessageBox.Show("Токен уже свободен.");
                return;
            }

            if (MessageBox.Show(
                $"Освободить токен {token.Sn}?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            var ok = await _tokenService.Unassign(token.Id);

            if (!ok)
            {
                MessageBox.Show(
                    "Ошибка освобождения токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            await LoadFromServer();
            AddToMiniLog($"Токен {token.Sn} освобожден");
        }

        private async void BtnDeleteToken_Click(object sender, RoutedEventArgs e)
        {
            if (_api == null)
            {
                MessageBox.Show("Сервер не подключен");
                return;
            }

            if (_tokenService == null)
            {
                MessageBox.Show("Сервис токенов не инициализирован");
                return;
            }

            if (!TryGetToken(sender, out var token))
                return;

            if (token.OwnerCertId != null)
            {
                MessageBox.Show(
                    "Нельзя удалить занятый токен.\nСначала освободите его.",
                    "Токен занят",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show(
                $"Удалить токен {token.Sn}?",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) != MessageBoxResult.Yes)
                return;

            var ok = await _tokenService.Delete(token.Id);

            if (!ok)
            {
                MessageBox.Show(
                    "Ошибка удаления токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                return;
            }

            if (_tokenService.Tokens.Any(t => t.Id == token.Id))
            {
                MessageBox.Show("Сервер ответил OK, но токен остался в списке.");
                return;
            }

            await LoadFromServer();
            AddToMiniLog($"Токен {token.Sn} удалён");
        }

        private bool TryGetToken(object sender, out TokenRecord token)
        {
            token = (sender as FrameworkElement)?.DataContext as TokenRecord;
            return token != null;
        }

        private void RebuildAvailableTokens()
        {
            if (_tokenService == null || _tokenService.Tokens == null)
                return;

            foreach (var cert in _allItems)
            {
                if (cert == null)
                    continue;

                if (cert.AvailableTokens == null)
                    cert.AvailableTokens = new ObservableCollection<TokenRecord>();

                var target = cert.AvailableTokens;
                target.Clear();

                foreach (var t in _tokenService.Tokens)
                {
                    if (t != null && (t.OwnerCertId == null || t.OwnerCertId == cert.Id))
                        target.Add(t);
                }
            }
        }

        private async void TokenComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_api == null)
                return;

            if (_tokenService == null)
                return;

            if (_isApplyingFromServer)
                return;

            if (_isAssigningToken)
                return;

            if (e.AddedItems.Count == 0)
                return;

            if (!(sender is ComboBox cb))
                return;

            if (!cb.IsKeyboardFocusWithin && !cb.IsDropDownOpen)
                return;

            if (!(cb.DataContext is CertRecord cert))
                return;

            var token = cb.SelectedItem as TokenRecord;
            if (token == null)
                return;

            if (cert.TokenId == token.Id)
                return;

            _isAssigningToken = true;

            try
            {
                var ok = await _tokenService.Assign(cert.Id, token.Id);

                if (!ok)
                {
                    MessageBox.Show(
                        "Ошибка назначения токена:\n" + (_tokenService.LastErrorMessage ?? "Нет ответа от сервера"),
                        "Ошибка",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);

                    await LoadFromServer();
                    return;
                }

                await LoadFromServer();
                AddToMiniLog($"Назначен токен {token.Sn} → {cert.Fio}");
            }
            catch (Exception ex)
            {
                await LoadFromServer();
                MessageBox.Show("Ошибка назначения токена:\n" + ex.Message);
            }
            finally
            {
                _isAssigningToken = false;
            }
        }
    }
}
using ImapCertWatcher.Client;
using ImapCertWatcher.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace ImapCertWatcher.Services
{
    public class TokenService
    {
        private readonly ServerApiClient _api;

        public ObservableCollection<TokenRecord> Tokens { get; } =
            new ObservableCollection<TokenRecord>();

        public event Action TokensChanged;

        public string LastErrorMessage { get; private set; }

        public TokenService(ServerApiClient api)
        {
            _api = api;
        }

        public async Task Reload()
        {
            var tokens = await _api.GetTokens() ?? new List<TokenRecord>();

            Tokens.Clear();
            foreach (var t in tokens)
                Tokens.Add(t);

            TokensChanged?.Invoke();
        }

        public async Task<bool> Assign(int certId, int tokenId)
        {
            LastErrorMessage = null;

            var resp = await _api.AssignToken(certId, tokenId);
            if (string.IsNullOrWhiteSpace(resp) || !resp.StartsWith("OK"))
            {
                LastErrorMessage = ExtractError(resp, "Ошибка назначения токена.");
                return false;
            }

            await Reload();
            return true;
        }

        public async Task<bool> Unassign(int tokenId)
        {
            LastErrorMessage = null;

            var resp = await _api.UnassignToken(tokenId);
            if (string.IsNullOrWhiteSpace(resp) || !resp.StartsWith("OK"))
            {
                LastErrorMessage = ExtractError(resp, "Ошибка освобождения токена.");
                return false;
            }

            await Reload();
            return true;
        }

        public async Task<bool> Delete(int tokenId)
        {
            LastErrorMessage = null;

            var resp = await _api.DeleteToken(tokenId);
            if (string.IsNullOrWhiteSpace(resp) || !resp.StartsWith("OK"))
            {
                LastErrorMessage = ExtractError(resp, "Ошибка удаления токена.");
                return false;
            }

            await Reload();
            return true;
        }

        private static string ExtractError(string response, string defaultText)
        {
            if (string.IsNullOrWhiteSpace(response))
                return "Нет ответа от сервера.";

            if (!response.StartsWith("ERROR"))
                return defaultText;

            int lastPipe = response.LastIndexOf('|');
            if (lastPipe >= 0 && lastPipe < response.Length - 1)
                return response.Substring(lastPipe + 1).Trim();

            return response.Trim();
        }
    }
}
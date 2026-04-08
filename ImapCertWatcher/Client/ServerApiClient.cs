using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public class ServerApiClient
    {
        private readonly ClientSettings _settings;

        public ServerApiClient(ClientSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public async Task<List<CertRecord>> GetCertificates()
        {
            return await GetListResponse<CertRecord>(
                "GET_CERTS",
                "CERTS ",
                "ERROR|GET_CERTS|",
                "GET_CERTS");
        }

        public async Task<List<TokenRecord>> GetTokens()
        {
            return await GetListResponse<TokenRecord>(
                "GET_TOKENS",
                "TOKENS ",
                "ERROR|GET_TOKENS|",
                "GET_TOKENS");
        }

        public async Task<List<TokenRecord>> GetFreeTokens()
        {
            return await GetListResponse<TokenRecord>(
                "GET_FREE_TOKENS",
                "TOKENS ",
                "ERROR|GET_FREE_TOKENS|",
                "GET_FREE_TOKENS");
        }

        public async Task<string> AssignToken(int certId, int tokenId)
        {
            return await SendCommandTrimmed($"SET_TOKEN|{certId}|{tokenId}");
        }

        public async Task<string> AddToken(string sn)
        {
            return await SendCommandTrimmed($"ADD_TOKEN|{sn}");
        }

        public async Task<string> DeleteToken(int id)
        {
            return await SendCommandTrimmed($"DELETE_TOKEN|{id}");
        }

        public async Task<string> UnassignToken(int id)
        {
            return await SendCommandTrimmed($"UNASSIGN_TOKEN|{id}");
        }

        public async Task<string> UpdateNote(int certId, string base64)
        {
            return await SendCommandTrimmed($"UPDATE_NOTE|{certId}|{base64}");
        }

        public async Task<string> AddArchive(int certId, string file, string base64)
        {
            return await SendCommandTrimmed($"ADD_ARCHIVE|{certId}|{file}|{base64}");
        }

        public async Task<string> GetArchive(int certId)
        {
            return await SendCommandTrimmed($"GET_ARCHIVE|{certId}");
        }

        public async Task<string> MarkDeleted(int id, bool deleted)
        {
            return await SendCommandTrimmed($"MARK_DELETED|{id}|{deleted}");
        }

        public async Task<string> UpdateBuilding(int id, string building)
        {
            building = building ?? "";
            return await SendCommandTrimmed($"SET_BUILDING|{id}|{building}");
        }

        public async Task<string> ResetRevokes()
        {
            return await SendCommandTrimmed("RESET_REVOKES");
        }

        public async Task<string> GetServerLog()
        {
            var response = await SendCommand("GET_LOG");

            if (string.IsNullOrWhiteSpace(response))
                return string.Empty;

            response = response.TrimEnd('\r', '\n');

            if (response.StartsWith("LOG "))
                return response.Substring(4);

            return response;
        }

        public async Task<string> SendCommand(string cmd)
        {
            return await TcpCommandClient.SendAsync(
                _settings.ServerIp,
                _settings.ServerPort,
                cmd);
        }

        private async Task<string> SendCommandTrimmed(string cmd)
        {
            var resp = await SendCommand(cmd);
            return resp?.Trim();
        }

        private async Task<List<T>> GetListResponse<T>(
            string command,
            string successPrefix,
            string errorPrefix,
            string operationName)
        {
            var response = await SendCommand(command);

            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("Пустой ответ сервера на " + operationName + ".");

            response = response.Trim();

            if (response.StartsWith(errorPrefix))
                throw new InvalidOperationException(response.Substring(errorPrefix.Length));

            if (!response.StartsWith(successPrefix))
                throw new InvalidOperationException(
                    "Некорректный ответ сервера на " + operationName + ": " + response);

            var json = response.Substring(successPrefix.Length);

            try
            {
                return JsonConvert.DeserializeObject<List<T>>(json) ?? new List<T>();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Ошибка разбора ответа " + operationName + ": " + ex.Message);
            }
        }
    }
}
using ImapCertWatcher.Models;
using ImapCertWatcher.Utils;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ImapCertWatcher.Client
{
    public class ServerApiClient
    {
        private readonly ClientSettings _settings;

        public ServerApiClient(ClientSettings settings)
        {
            _settings = settings;
        }

        public async Task<List<CertRecord>> GetCertificates()
        {
            var response = await SendCommand("GET_CERTS");

            if (string.IsNullOrWhiteSpace(response))
                return new List<CertRecord>();

            response = response.Trim();

            if (response.StartsWith("CERTS "))
                response = response.Substring(6);

            try
            {
                return JsonConvert.DeserializeObject<List<CertRecord>>(response)
                       ?? new List<CertRecord>();
            }
            catch
            {
                return new List<CertRecord>();
            }
        }

        public async Task<List<TokenRecord>> GetTokens()
        {
            var response = await SendCommand("GET_TOKENS");

            if (string.IsNullOrWhiteSpace(response))
                return new List<TokenRecord>();

            response = response.Trim();

            if (response.StartsWith("TOKENS "))
                response = response.Substring(7);

            try
            {
                return JsonConvert.DeserializeObject<List<TokenRecord>>(response)
                       ?? new List<TokenRecord>();
            }
            catch
            {
                return new List<TokenRecord>();
            }
        }

        public async Task<List<TokenRecord>> GetFreeTokens()
        {
            var response = await SendCommand("GET_FREE_TOKENS");

            if (string.IsNullOrWhiteSpace(response))
                return new List<TokenRecord>();

            response = response.Trim();

            if (response.StartsWith("TOKENS "))
                response = response.Substring(7);

            try
            {
                return JsonConvert.DeserializeObject<List<TokenRecord>>(response)
                       ?? new List<TokenRecord>();
            }
            catch
            {
                return new List<TokenRecord>();
            }
        }

        public async Task<string> AssignToken(int certId, int tokenId)
        {
            var resp = await SendCommand($"SET_TOKEN|{certId}|{tokenId}");
            return resp?.Trim();
        }

        public async Task<string> AddToken(string sn)
        {
            var resp = await SendCommand($"ADD_TOKEN|{sn}");
            return resp?.Trim();
        }

        public async Task<string> DeleteToken(int id)
        {
            var resp = await SendCommand($"DELETE_TOKEN|{id}");
            return resp?.Trim();
        }

        public async Task<string> UnassignToken(int id)
        {
            var resp = await SendCommand($"UNASSIGN_TOKEN|{id}");
            return resp?.Trim();
        }

        public async Task<string> UpdateNote(int certId, string base64)
        {
            return await SendCommand($"UPDATE_NOTE|{certId}|{base64}");
        }

        public async Task<string> AddArchive(int certId, string file, string base64)
        {
            return await SendCommand($"ADD_ARCHIVE|{certId}|{file}|{base64}");
        }

        public async Task<string> GetArchive(int certId)
        {
            return await SendCommand($"GET_ARCHIVE|{certId}");
        }

        public async Task<string> SendCommand(string cmd)
        {
            return await TcpCommandClient.SendAsync(
                _settings.ServerIp,
                _settings.ServerPort,
                cmd);
        }

        public async Task<string> MarkDeleted(int id, bool deleted)
        {
            return await SendCommand($"MARK_DELETED|{id}|{deleted}");
        }

        public async Task<string> UpdateBuilding(int id, string building)
        {
            building = building ?? "";
            return await SendCommand($"SET_BUILDING|{id}|{building}");
        }

        public async Task<string> ResetRevokes()
        {
            return await SendCommand("RESET_REVOKES");
        }
    }
}

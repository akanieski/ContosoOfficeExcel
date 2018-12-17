using Microsoft.AspNetCore.Authentication.AzureAD.UI;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace ContosoWebPortal
{
    public interface IGraphService
    {
        GraphServiceClient GetAuthenticatedClient();
        Task<DriveItem> UploadFileToUsersDocuments(string userId, string sourceFile, string targetPath = "");
        Task<DriveItem> UploadFileToUsersDocuments(string userId, string fileName, Stream sourceStream, string targetPath = "");
        Task<DriveItem> UploadFileToDrive(string driveId, string sourceFile, string targetPath = "");
        Task<IEnumerable<Drive>> GetDrivesByUser(string userId);
        Task<Drive> GetUsersDocumentsDrive(string userId);
    }
    public class GraphService : IGraphService
    {
        public GraphService(IConfiguration config)
        {
            config.Bind("AzureAd", _azureADOptions = new CustomAzureADOptions());
            _http = new HttpClient();
        }
        private const string _rootUrl = "https://graph.microsoft.com/v1.0";
        private TokenResponse _token;
        private CustomAzureADOptions _azureADOptions;
        private HttpClient _http;
        private DateTime _lastTimeTokenReceived;
        private GraphServiceClient _graphClient;

        public GraphServiceClient GetAuthenticatedClient()
        {
            if (_graphClient == null)
            {
                // Create Microsoft Graph client.
                try
                {
                    _graphClient = new GraphServiceClient(
                        "https://graph.microsoft.com/v1.0",
                        new DelegateAuthenticationProvider(
                            async (requestMessage) =>
                            {
                                await GetToken();
                                requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token.AccessToken);
                                // This header has been added to identify our sample in the Microsoft Graph service.  If extracting this code for your project please remove.


                            }));
                    return _graphClient;
                }

                catch (Exception ex)
                {
                    Debug.WriteLine("Could not create a graph client: " + ex.Message);
                }
            }

            return _graphClient;
        }

        public async Task GetToken()
        {
            var req = new HttpRequestMessage(HttpMethod.Post, $"https://login.microsoftonline.com/{_azureADOptions.TenantId}/oauth2/v2.0/token/");
            req.Content = new FormUrlEncodedContent(new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_secret", _azureADOptions.ClientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
                new KeyValuePair<string, string>("client_id", _azureADOptions.ClientId),
            });
            var res = await _http.SendAsync(req);
            var stringData = await res.Content.ReadAsStringAsync();
            var data = JsonConvert.DeserializeObject<TokenResponse>(stringData);
            _token = data;
            _lastTimeTokenReceived = DateTime.Now;
        }

        /// <summary>
        /// Uploads a file to a specific drive.
        /// </summary>
        /// <param name="driveId">Drive Object Id</param>
        /// <param name="sourceFile">File to upload on local system</param>
        /// <param name="targetPath">[Optional] folder path, ie. "contoso/docs/" </param>
        /// <returns></returns>
        public async Task<DriveItem> UploadFileToDrive(string driveId, string sourceFile, string targetPath = "")
        {
            return await GetAuthenticatedClient()
                .Drives[driveId].Root
                .ItemWithPath(targetPath + Path.GetFileName(sourceFile)).Content.Request()
                .PutAsync<DriveItem>(new FileStream(sourceFile, FileMode.Open));
        }

        /// <summary>
        /// Uploads a file to a specific drive.
        /// </summary>
        /// <param name="userId">User Object Id</param>
        /// <param name="sourceFile">File to upload on local system</param>
        /// <param name="targetPath">[Optional] folder path, ie. "contoso/docs/" </param>
        /// <returns></returns>
        public async Task<DriveItem> UploadFileToUsersDocuments(string userId, string sourceFile, string targetPath = "")
        {
            return await GetAuthenticatedClient()
                .Users[userId].Drive.Root
                .ItemWithPath(targetPath + Path.GetFileName(sourceFile)).Content.Request()
                .PutAsync<DriveItem>(new FileStream(sourceFile, FileMode.Open));
        }

        /// <summary>
        /// Uploads a file to a specific drive.
        /// </summary>
        /// <param name="userId">User Object Id</param>
        /// <param name="sourceFile">File to upload on local system</param>
        /// <param name="targetPath">[Optional] folder path, ie. "contoso/docs/" </param>
        /// <returns></returns>
        public async Task<DriveItem> UploadFileToUsersDocuments(string userId, string fileName, Stream sourceStream, string targetPath = "")
        {
            return await GetAuthenticatedClient()
                .Users[userId].Drive.Root
                .ItemWithPath(targetPath + fileName).Content.Request()
                .PutAsync<DriveItem>(sourceStream);
        }

        public async Task<IEnumerable<Drive>> GetDrivesByUser(string userId)
        {
            return await GetAuthenticatedClient()
                .Users[userId].Drives.Request().GetAsync();
        }

        public async Task<Drive> GetUsersDocumentsDrive(string userId)
        {
            return await GetAuthenticatedClient()
                .Users[userId].Drive.Request().GetAsync();
        }

        /*
        private async Task EnsureToken()
        {
            if (_token == null || _lastTimeTokenReceived == null || (DateTime.Now - _lastTimeTokenReceived).TotalSeconds >= _token.ExpiresIn - 1)
            {
                await GetToken();
            }
        }

        public async Task<DriveDetails> GetDrive(string userObjectId)
        {
            await EnsureToken();
            var req = new HttpRequestMessage(HttpMethod.Get, $"{_rootUrl}/users/{userObjectId}/drive");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer",_token.AccessToken);
            var res = await _http.SendAsync(req);
            var stringData = await res.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<DriveDetails>(stringData);
        }

        public async Task UploadFile(string driveId, string sourceFilePath, string targetFolder)
        {
            await EnsureToken();
            await CreateFolder(driveId, targetFolder);
            var url = $"{_rootUrl}/drives/{driveId}/items/{targetFolder}:/{System.IO.Path.GetFileName(sourceFilePath)}:/content";
            var data = System.IO.File.ReadAllBytes(sourceFilePath);
            var req = new HttpRequestMessage(HttpMethod.Put, url);
            req.Content = new ByteArrayContent(data);
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _token.AccessToken);
            var resp = await _http.SendAsync(req);
            var stringData = resp.Content.ReadAsStringAsync();
        }

        public async Task CreateFolder(string driveId, string folderName)
        {
            await EnsureToken();
            var url = $"{_rootUrl}/drives/{driveId}/root/children";
            var data = new CreateFolderRequest
            {
                Name = folderName,
                Folder = new { },
                ConflictBehavior = "rename"
            };
            var req = new HttpRequestMessage(HttpMethod.Post, url);
            req.Content = new StringContent(JsonConvert.SerializeObject(data));
            req.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");
            req.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", _token.AccessToken);
            var resp = await _http.SendAsync(req);
            var stringData = resp.Content.ReadAsStringAsync();
        }
        */
    }


    public class CreateFolderRequest
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("folder")]
        public object Folder{ get; set; }
        [JsonProperty("@microsoft.graph.conflictBehavior")]
        public string ConflictBehavior { get; set; }

    }
    public class DriveDetails
    {
        [JsonProperty("description")]
        public string Description { get; set; }
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("web_url")]
        public string WebUrl { get; set; }
    }

    public class CustomAzureADOptions : AzureADOptions
    {
        public string ClientSecret { get; set; }
    }

    public class TokenResponse
    {
        [JsonProperty("token_type")]
        public string TokenType { get; set; }
        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }
        [JsonProperty("ext_expires_in")]
        public int ExtExpiresIn { get; set; }
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
    
}

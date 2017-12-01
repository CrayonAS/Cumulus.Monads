using System;
using System.Collections.Generic;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Pzl.O365.ProvisioningFunctions.Helpers
{
    class ConnectADAL
    {
        private static readonly Uri ADALLogin = new Uri("https://login.windows.net/");
        const string GraphResourceId = "https://graph.microsoft.com"; // Microsoft Graph End-point
        private static string _appId = Environment.GetEnvironmentVariable("ADALAppId");
        private static string _appSecret = Environment.GetEnvironmentVariable("ADALAppSecret");
        private static string _appCert = Environment.GetEnvironmentVariable("ADALAppCertificate");
        private static string _appCertKey = Environment.GetEnvironmentVariable("ADALAppCertificateKey");
        private static readonly string ADALDomain = Environment.GetEnvironmentVariable("ADALDomain");
        private static readonly Dictionary<string, AuthenticationResult> ResourceTokenLookup = new Dictionary<string, AuthenticationResult>();
        private static readonly string MsiEndpoint = Environment.GetEnvironmentVariable("MSI_ENDPOINT");
        private static readonly string MsiSecret = Environment.GetEnvironmentVariable("MSI_SECRET");

        private static readonly AzureServiceTokenProvider AzureServiceTokenProvider = new AzureServiceTokenProvider();
        private static readonly KeyVaultClient KvClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(AzureServiceTokenProvider.KeyVaultTokenCallback));

        public class MsiInformation
        {
            public string OwnerId { get; set; }
            public string BearerToken { get; set; }
        }

        private static string SecretUri(string secret)
        {
            return $"{ConfigurationManager.AppSettings["KeyVaultUri"].TrimEnd('/')}/Secrets/{secret}";
        }

        private static async Task GetVariables()
        {
            //http://integration.team/2017/09/25/retrieve-azure-key-vault-secrets-using-azure-functions-managed-service-identity/
            if (string.IsNullOrEmpty(_appId)) _appId = (await KvClient.GetSecretAsync(SecretUri("ADALAppId"))).Value;
            if (string.IsNullOrEmpty(_appSecret)) _appSecret = (await KvClient.GetSecretAsync(SecretUri("ADALAppSecret"))).Value;
            if (string.IsNullOrEmpty(_appCert)) _appCert = (await KvClient.GetSecretAsync(SecretUri("ADALAppCertificate"))).Value;
            if (string.IsNullOrEmpty(_appCertKey)) _appCertKey = (await KvClient.GetSecretAsync(SecretUri("ADALAppCertificateKey"))).Value;
        }

        private static async Task<string> GetAccessToken(string AADDomain)
        {
            await GetVariables();
            AuthenticationResult token;
            if (ResourceTokenLookup.TryGetValue(GraphResourceId, out token) &&
                token.ExpiresOn.UtcDateTime >= DateTime.UtcNow) return token.AccessToken;

            var authenticationContext = new AuthenticationContext(ADALLogin + AADDomain);
            var clientCredential = new ClientCredential(_appId, _appSecret);

            bool keepRetry = false;
            do
            {
                TimeSpan? delay = null;
                try
                {
                    token = await authenticationContext.AcquireTokenAsync(GraphResourceId, clientCredential);
                }
                catch (Exception ex)
                {
                    if (!(ex is AdalServiceException) && !(ex.InnerException is AdalServiceException)) throw;

                    AdalServiceException serviceException;
                    if (ex is AdalServiceException) serviceException = (AdalServiceException)ex;
                    else serviceException = (AdalServiceException)ex.InnerException;
                    if (serviceException.ErrorCode == "temporarily_unavailable")
                    {
                        RetryConditionHeaderValue retry = serviceException.Headers.RetryAfter;
                        if (retry.Delta.HasValue)
                        {
                            delay = retry.Delta;
                        }
                        else if (retry.Date.HasValue)
                        {
                            delay = retry.Date.Value.Offset;
                        }
                        if (delay.HasValue)
                        {
                            Thread.Sleep((int)delay.Value.TotalSeconds); // sleep or other
                            keepRetry = true;
                        }
                    }
                    else
                    {
                        throw;
                    }
                }
            } while (keepRetry);

            ResourceTokenLookup[GraphResourceId] = token;
            return token.AccessToken;
        }

        private static async Task<string> GetAccessTokenSharePoint(string AADDomain, string siteUrl, TraceWriter log = null)
        {
            //https://blogs.msdn.microsoft.com/richard_dizeregas_blog/2015/05/03/performing-app-only-operations-on-sharepoint-online-through-azure-ad/
            AuthenticationResult token;
            Uri uri = new Uri(siteUrl);
            string resourceUri = uri.Scheme + "://" + uri.Authority;
            if (!ResourceTokenLookup.TryGetValue(resourceUri, out token) || token.ExpiresOn.UtcDateTime < DateTime.UtcNow)
            {
                if (token != null)
                {
                    log?.Info($"Token expired {token.ExpiresOn.UtcDateTime}");
                }

                var cac = GetClientAssertionCertificate();
                var authenticationContext = new AuthenticationContext(ADALLogin + AADDomain);

                bool keepRetry = false;
                do
                {
                    TimeSpan? delay = null;
                    try
                    {
                        token = await authenticationContext.AcquireTokenAsync(resourceUri, cac);
                    }
                    catch (Exception ex)
                    {
                        if (!(ex is AdalServiceException) && !(ex.InnerException is AdalServiceException)) throw;

                        AdalServiceException serviceException;
                        if (ex is AdalServiceException) serviceException = (AdalServiceException)ex;
                        else serviceException = (AdalServiceException)ex.InnerException;
                        if (serviceException.ErrorCode == "temporarily_unavailable")
                        {
                            RetryConditionHeaderValue retry = serviceException.Headers.RetryAfter;
                            if (retry.Delta.HasValue)
                            {
                                delay = retry.Delta;
                            }
                            else if (retry.Date.HasValue)
                            {
                                delay = retry.Date.Value.Offset;
                            }
                            if (delay.HasValue)
                            {
                                Thread.Sleep((int)delay.Value.TotalSeconds); // sleep or other
                                keepRetry = true;
                            }
                        }
                        else
                        {
                            throw;
                        }
                    }
                } while (keepRetry);

                //token = await authenticationContext.AcquireTokenAsync(resourceUri, cac);
                ResourceTokenLookup[resourceUri] = token;

                log?.Info($"Aquired token which expires {token.ExpiresOn.UtcDateTime}");

            }
            return token.AccessToken;
        }

        public static GraphServiceClient GetGraphClient()
        {
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    string accessToken = await GetAccessToken(ADALDomain);
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
            return client;
        }

        public static GraphServiceClient GetGraphClientServiceIdentity(TraceWriter log)
        {
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    var info = await GetBearerTokenServiceIdentity(log);
                    log.Info("Bearer: " + info.BearerToken);
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", info.BearerToken);
                }));
            return client;
        }

        public static async Task<string> GetBearerToken()
        {
            string accessToken = await GetAccessToken(ADALDomain);
            return accessToken;
        }

        public static async Task<MsiInformation> GetBearerTokenServiceIdentity(TraceWriter log)
        {
            string apiVersion = "2017-09-01";
            string tokenAuthUri = MsiEndpoint + $"?resource={GraphResourceId}&api-version={apiVersion}";
            log.Info(tokenAuthUri);
            HttpClient client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.TryAddWithoutValidation("Secret", MsiSecret);

            var response = await client.GetAsync(tokenAuthUri);
            if (response.IsSuccessStatusCode)
            {
                string responseMsg = await response.Content.ReadAsStringAsync();
                dynamic token = JsonConvert.DeserializeObject(responseMsg);
                string bearer = token.access_token;

                var parts = bearer.Split('.');
                var decoded = Convert.FromBase64String(parts[1]);
                var part = Encoding.UTF8.GetString(decoded);
                var jwt = JObject.Parse(part);
                var ownerId = jwt["oid"].Value<string>();

                MsiInformation info = new MsiInformation
                {
                    OwnerId = ownerId,
                    BearerToken = bearer
                };
                log.Info("Got token: " + bearer);
                return info;
            }
            log.Info("No token");
            return null;
        }

        private static ClientAssertionCertificate GetClientAssertionCertificate()
        {
            var generator = new Certificate.Certificate(_appCert, _appCertKey, "");
            X509Certificate2 cert = generator.GetCertificateFromPEMstring(false);
            ClientAssertionCertificate cac = new ClientAssertionCertificate(_appId, cert);
            return cac;
        }

        public static async Task<ClientContext> GetClientContext(string url, TraceWriter log = null)
        {
            string bearerToken = await GetAccessTokenSharePoint(ADALDomain, url, log);
            var clientContext = new ClientContext(url);
            clientContext.ExecutingWebRequest += (sender, args) =>
            {
                args.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + bearerToken;
            };
            return clientContext;
        }

    }
}

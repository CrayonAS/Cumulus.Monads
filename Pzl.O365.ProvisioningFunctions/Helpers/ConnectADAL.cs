using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using ADAL = Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;

namespace Pzl.O365.ProvisioningFunctions.Helpers
{
    class ConnectADAL
    {
        private static readonly Uri ADALLogin = new Uri("https://login.windows.net/");
        const string GraphResourceId = "https://graph.microsoft.com"; // Microsoft Graph End-point
        internal static readonly string AppId = Environment.GetEnvironmentVariable("ADALAppId");
        private static readonly string AppSecret = Environment.GetEnvironmentVariable("ADALAppSecret");
        private static readonly string AppCert = Environment.GetEnvironmentVariable("ADALAppCertificate");
        private static readonly string AppCertKey = Environment.GetEnvironmentVariable("ADALAppCertificateKey");
        private static readonly string ADALDomain = Environment.GetEnvironmentVariable("ADALDomain");
        private static readonly Dictionary<string, ADAL.AuthenticationResult> ResourceTokenLookup = new Dictionary<string, ADAL.AuthenticationResult>();


        private static async Task<string> GetAccessToken(string AADDomain)
        {
            ADAL.AuthenticationResult token;
            if (!ResourceTokenLookup.TryGetValue(GraphResourceId, out token) || token.ExpiresOn.UtcDateTime < DateTime.UtcNow)
            {
                var authenticationContext = new ADAL.AuthenticationContext(ADALLogin + AADDomain);
                var clientCredential = new ADAL.ClientCredential(AppId, AppSecret);

                //var url = await authenticationContext.GetAuthorizationRequestUrlAsync(resourceUri, AppId, new Uri("https://techmikael.sharepoint.com/o365"), ADAL.UserIdentifier.AnyUser, "prompt=admin_consent");

                token = await authenticationContext.AcquireTokenAsync(GraphResourceId, clientCredential);
                ResourceTokenLookup[GraphResourceId] = token;
            }
            return token.AccessToken;
        }

        private static async Task<string> GetAccessTokenSharePoint(string AADDomain, string siteUrl, TraceWriter log = null)
        {
            //https://blogs.msdn.microsoft.com/richard_dizeregas_blog/2015/05/03/performing-app-only-operations-on-sharepoint-online-through-azure-ad/
            ADAL.AuthenticationResult token;
            Uri uri = new Uri(siteUrl);
            string resourceUri = uri.Scheme + "://" + uri.Authority;
            if (!ResourceTokenLookup.TryGetValue(resourceUri, out token) || token.ExpiresOn.UtcDateTime < DateTime.UtcNow)
            {
                if (token != null)
                {
                    log?.Info($"Token expired {token.ExpiresOn.UtcDateTime}");
                }

                var cac = GetClientAssertionCertificate();
                var authenticationContext = new ADAL.AuthenticationContext(ADALLogin + AADDomain);
                token = await authenticationContext.AcquireTokenAsync(resourceUri, cac);
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

        public static async Task<string> GetBearerToken()
        {
            string accessToken = await GetAccessToken(ADALDomain);
            return accessToken;
        }



        private static ADAL.ClientAssertionCertificate GetClientAssertionCertificate()
        {
            var generator = new Certificate.Certificate(AppCert, AppCertKey, "");
            X509Certificate2 cert = generator.GetCertificateFromPEMstring(false);
            ADAL.ClientAssertionCertificate cac = new ADAL.ClientAssertionCertificate(AppId, cert);
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

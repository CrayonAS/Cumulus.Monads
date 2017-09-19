using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;
using Newtonsoft.Json;
using ADAL = Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace CreateSchemaConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            DoWork().Wait();
        }

        private const string MetadataSchemaName = "CollaborationMetadata";
        static async Task DoWork()
        {
            var client = GetGraphClient();
            var res = await client.SchemaExtensions.Request().GetAsync();
            SchemaExtension gm = null;
            foreach (SchemaExtension schemaExtension in res)
            {
                //Console.WriteLine(schemaExtension.Id);
                if (schemaExtension.Id.EndsWith(MetadataSchemaName))
                {
                    gm = schemaExtension;
                    break;
                }
            }

            if (gm != null)
            {
                //gm.Status = "Available";
                //await client.SchemaExtensions[gm.Id].Request().UpdateAsync(gm);
                Console.WriteLine(JsonConvert.SerializeObject(gm));
                ////gm.Status = "InDevelopment";
                //await client.SchemaExtensions[gm.Id].Request().UpdateAsync(gm);

                //await client.SchemaExtensions[gm.Id].Request().DeleteAsync();

                //ExtensionSchemaProperty groupTypeProperty = new ExtensionSchemaProperty
                //{
                //    Name = "groupType",
                //    Type = "String"
                //};
                //ExtensionSchemaProperty responsibleProperty = new ExtensionSchemaProperty
                //{
                //    Name = "responsible",
                //    Type = "String"
                //};
                //gm.Properties = new[] { groupTypeProperty, responsibleProperty };
                //await client.SchemaExtensions[gm.Id].Request().UpdateAsync(gm);
                //Console.WriteLine("patched");
            }

            // Create schema
            //await CreateSchema(client);
        }

        private static async Task CreateSchema(GraphServiceClient client)
        {
            SchemaExtension groupMetadata = new SchemaExtension
            {
                Id = MetadataSchemaName,
                Description = "Schema describing extra metadata for an Office 365 Group",
                TargetTypes = new[] {"Group"},
                Owner = "4fdcefb4-194a-4412-8ca7-07a27a830248"
                //,Status = "Available"
            };
            ExtensionSchemaProperty groupTypeProperty = new ExtensionSchemaProperty
            {
                Name = "GroupType",
                Type = "String"
            };
            ExtensionSchemaProperty responsibleProperty = new ExtensionSchemaProperty
            {
                Name = "Responsible",
                Type = "String"
            };
            ExtensionSchemaProperty startDateProperty = new ExtensionSchemaProperty
            {
                Name = "StartDate",
                Type = "DateTime"
            };
            ExtensionSchemaProperty endDateProperty = new ExtensionSchemaProperty
            {
                Name = "EndDate",
                Type = "DateTime"
            };
            groupMetadata.Properties = new[] {groupTypeProperty, responsibleProperty, startDateProperty, endDateProperty};
            var createdExtension = await client.SchemaExtensions.Request().AddAsync(groupMetadata);
            var extensionName = createdExtension.Id;
        }

        public static GraphServiceClient GetGraphClient()
        {
            GraphServiceClient client = new GraphServiceClient(new DelegateAuthenticationProvider(
                async (requestMessage) =>
                {
                    string accessToken = await GetToken();
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                }));
            return client;
        }

        static async Task<string> GetToken()
        {
            // The O365 login URL
            string authorityUri = "https://login.windows.net/common/oauth2/authorize";

            const string clientId = "4fdcefb4-194a-4412-8ca7-07a27a830248";
            

            // Replu URL for the Microsoft SharePoint Online Management Shell
            const string redirectUri = "https://techmikael.sharepoint.com/o365";

            // The DOMAIN you want to do API calls against
            string resourceUri = "https://graph.microsoft.com";

            IntPtr handle = Process.GetCurrentProcess().MainWindowHandle;

            var username = "miksvenson@techmikael.onmicrosoft.com";
            var pw = "<>";
            var secret = "<>";
            var url = "https://login.microsoftonline.com/techmikael.onmicrosoft.com/oauth2/token";
            var payload =
                $"resource={resourceUri}&client_id={clientId}&grant_type=password&username={username}&password={pw}&scope=openid&client_secret={secret}";
            //var payload =
            //    $"resource={resourceUri}&client_id={clientId}&grant_type=password&username={username}&password={pw}&scope=openid";
            var content = new StringContent(payload, Encoding.UTF8, "application/x-www-form-urlencoded");

            HttpClient client = new HttpClient();
            var response = await client.PostAsync(url, content);
            string responseMsg = await response.Content.ReadAsStringAsync();
            dynamic auth = JsonConvert.DeserializeObject(responseMsg);
            return auth.access_token;

            var authCtx = new ADAL.AuthenticationContext(authorityUri);
            var authParam = new ADAL.PlatformParameters(ADAL.PromptBehavior.Auto, handle);
            var authenticationResult = await authCtx.AcquireTokenAsync(resourceUri, clientId, new Uri(redirectUri), authParam, ADAL.UserIdentifier.AnyUser, $"client_secret={secret}");

            return authenticationResult.AccessToken;
        }
    }
}

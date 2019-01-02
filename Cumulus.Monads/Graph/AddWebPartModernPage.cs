using System.Linq;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using ClientSidePage = OfficeDevPnP.Core.Pages.ClientSidePage;
using System.Threading;
using System.Collections.Generic;
using OfficeDevPnP.Core.Pages;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace AddModernWebpartToPage
{
    public static class AddWebPartModernPage
    {
        [FunctionName("AddWebPartModernPage")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {


            // These can be passed as parameters and some other part properties.
            string webUrl = "https://m365x456568.sharepoint.com";
            string userName = "admin@M365x456568.onmicrosoft.com";
            string pwd = "Dhaahgaab1";
            
            // I used this because the ADAL connection did work for me.
            var password = new SecureString();
            foreach (var c in pwd.ToCharArray()) { password.AppendChar(c); }
            using (var ctx = new ClientContext(webUrl))
            {
                var web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.Credentials = new SharePointOnlineCredentials(userName, password);
                ctx.RequestTimeout = Timeout.Infinite;
                ctx.ExecuteQuery();

                string pageName = "FirstTest.aspx";

                ClientSidePage page = ClientSidePage.Load(ctx, pageName);

                var components = page.AvailableClientSideComponents();
                string wpName = "0ef418ba-5d19-4ade-9db0-b339873291d0";
                List<string> componentsnames = components.Select(k => k.Name).ToList();
                var webPartToAdd = components.Where(wp => wp.ComponentType == 1 && wp.Name == wpName).FirstOrDefault();
                webPartToAdd.Manifest = "{\"$schema\": \"https://dev.office.com/json-schemas/spfx/client-side-web-part-manifest.schema.json\",\"id\": \"0ef418ba-5d19-4ade-9db0-b339873291d0\",\"alias\": \"MyInformationEventsWebPart\",\"componentType\": \"WebPart\",\"version\": \"*\",\"manifestVersion\": 2,\"requiresCustomScript\": false,\"preconfiguredEntries\": [{\"groupId\": \"5c03119e-3074-46fd-976b-c60198311f70\", \"group\": { \"default\": \"Other\" },\"title\": { \"default\": \"My Information and Events\" },\"description\": { \"default\": \"Lists the information and events\" },\"officeFabricIconFontName\": \"Sunny\",\"properties\": {  \"description\": \"myInformationEvents\"}}]}";

                if (webPartToAdd != null)
                {
                    ClientSideWebPart clientWp = new ClientSideWebPart(webPartToAdd) { Order = -1, Title = "Cumulus", Description = "Automation Engine Web Parts" };
                    clientWp.PropertiesJson = "{\"description\":\"myInformationEvents\",\"layout\":\"1\",\"enableGlobalContent\":true,\"enableBUSiteOnlyFiltering\":true,\"cardViewSelection\":\"single\"}";

                    page.AddControl(clientWp);
                }
                page.Save(pageName);
                page.Publish();
                ctx.ExecuteQuery();

                return null;

            }
        }
    }
}









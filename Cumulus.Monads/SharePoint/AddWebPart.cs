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
using System.ComponentModel.DataAnnotations;
using Cumulus.Monads.Helpers;
using Group = Microsoft.Graph.Group;
using Newtonsoft.Json;
using System.Net.Http.Formatting;
using System.Web.UI.WebControls;
using System.Net;
using System.Web;

namespace Cumulus.Monads.Graph
{
    public static class AddWebPart
    {
        [FunctionName("AddWebPart")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]AddWebPartRequest req, TraceWriter log)
        {
            
            string webUrl = req.SiteURL;
            string pageName = req.PageName;
            var ctx = await ConnectADAL.GetClientContext(req.SiteURL, log);
            var web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.RequestTimeout = Timeout.Infinite;
            ctx.ExecuteQuery();

            ClientSidePage page = ClientSidePage.Load(ctx, pageName);

            var components = page.AvailableClientSideComponents();
            List<string> componentsnames = components.Select(k => k.Name).ToList();
            var webPartToAdd = components.Where(wp => wp.ComponentType == 1 && wp.Name == req.WebPartName).FirstOrDefault();

            if (webPartToAdd != null)
            {
                ClientSideWebPart clientWp = new ClientSideWebPart(webPartToAdd) {  Order = req.Order,  Title = req.Title, Description = req.Description, PropertiesJson = req.PropertiesJson};
                page.AddControl(clientWp);
            }

            page.Save(pageName);
            page.Publish();
            ctx.ExecuteQuery();
                       
            var CreateWebPartResponse = new AddWebPartResponse
            {
                
                SiteURL = req.SiteURL
            };

            return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
            {
                Content = new ObjectContent<AddWebPartResponse>(CreateWebPartResponse, new JsonMediaTypeFormatter())
            });

        }

        public class AddWebPartRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Name of the Page")]
            public string PageName { get; set; }

            [Required]
            [Display(Description = "Name of the Web Part or ID")]
            public string WebPartName { get; set; }

            [Required]
            [Display(Description = "Controll Order")]
            public int Order { get; set; }

            [Display(Description = "Description of the Web Part")]
            public string Description { get; set; }

            [Display(Description = "Title of the Web Part")]
            public string Title { get; set; }

            [Display(Description = "Web Properties")]
            public string Properties { get;}

            [Display(Description = "Web Part Information")]
            public string PropertiesJson { get; set; }

            [Display(Description = "Server Processed Content")]
            public string ServerProcessedContent { get; }
        }
       
        public class AddWebPartResponse
        {
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }

        }
    }
}

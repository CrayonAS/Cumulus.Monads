using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Enums;
using Pzl.O365.ProvisioningFunctions.Helpers;

namespace Pzl.O365.ProvisioningFunctions.SharePoint
{
    public static class GetTermProperty
    {
        [FunctionName("GetTermProperty")]
        [ResponseType(typeof(GetTermPropertyResponse))]
        [Display(Name = "Get term property", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]GetTermPropertyRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);
                clientContext.Load(taxonomySession);
                clientContext.ExecuteQuery();
                TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();                 
                var term = termStore.GetTerm(request.TermGUID);
                clientContext.Load(term, t => t.LocalCustomProperties);
                clientContext.ExecuteQuery();
                var propertyValue = term.LocalCustomProperties[request.PropertyName];
                var getTermPropertyResponse = new GetTermPropertyResponse
                {
                    PropertyValue = propertyValue
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<GetTermPropertyResponse>(getTermPropertyResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class GetTermPropertyRequest
        {
            [Required]
            public string SiteURL { get; set; }

            [Required]
            [Display(Description = "Term GUID")]
            public Guid TermGUID { get; set; }

            [Required]
            [Display(Description = "Property name")]
            public string PropertyName { get; set; }
        }

        public class GetTermPropertyResponse
        {
            public string PropertyValue { get; set; }
        }
    }
}

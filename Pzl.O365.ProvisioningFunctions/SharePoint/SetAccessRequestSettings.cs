using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Pzl.O365.ProvisioningFunctions.Helpers;

namespace Pzl.O365.ProvisioningFunctions.SharePoint
{
    public static class SetAccessRequestSettings
    {
        [FunctionName("SetAccessRequestSettings")]
        [ResponseType(typeof(SetAccessRequestSettingsResponse))]
        [Display(Name = "Set access request settings", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetAccessRequestSettingsRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;
            bool isDirty = false;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);

                var web = clientContext.Web;
                clientContext.Load(web, w => w.MembersCanShare, w => w.AssociatedMemberGroup.AllowMembersEditMembership, w => w.RequestAccessEmail);
                clientContext.ExecuteQuery();


                if (request.MembersCanShare != web.MembersCanShare)
                {
                    isDirty = true;
                }
                web.MembersCanShare = request.MembersCanShare;
                web.Update();

                if (request.AllowMembersEditMembership != web.AssociatedMemberGroup.AllowMembersEditMembership)
                {
                    isDirty = true;
                }
                web.AssociatedMemberGroup.AllowMembersEditMembership = request.AllowMembersEditMembership;
                web.AssociatedMemberGroup.Update();

                if (request.RequestAccessEmail != null)
                {

                    if (request.RequestAccessEmail != web.RequestAccessEmail)
                    {
                        isDirty = true;
                    }
                    web.RequestAccessEmail = request.RequestAccessEmail;
                    web.Update();
                }
                clientContext.ExecuteQuery();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetAccessRequestSettingsResponse>(new SetAccessRequestSettingsResponse { AccessRequestSettingsModified = isDirty }, new JsonMediaTypeFormatter())
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

        public class SetAccessRequestSettingsRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
            [Display(Description = "Allow members to share the site and individual files and folders.")]
            public bool MembersCanShare { get; set; }
            [Display(Description = "Send all requests for access to the following e-mail address")]
            public string RequestAccessEmail { get; set; }
            [Display(Description = "Allow members to invite others to the site members group. This setting must be enabled to let members share the site.")]
            public bool AllowMembersEditMembership { get; set; }
        }

        public class SetAccessRequestSettingsResponse
        {
            [Display(Description = "")]
            public bool AccessRequestSettingsModified { get; set; }
        }
    }
}

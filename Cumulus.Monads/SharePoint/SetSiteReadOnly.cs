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
using Cumulus.Monads.Helpers;
using Microsoft.SharePoint.Client.InformationPolicy;

namespace Cumulus.Monads.SharePoint
{
    public static class SetSiteReadOnly
    {
        [FunctionName("SetSiteReadOnly")]
        [ResponseType(typeof(SetSiteReadOnlyResponse))]
        [Display(Name = "Set the site to read-only", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteReadOnlyRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                if (string.IsNullOrWhiteSpace(request.SiteURL))
                {
                    throw new ArgumentException("Parameter cannot be null", "SiteURL");
                }
                if (string.IsNullOrWhiteSpace(request.Owner))
                {
                    throw new ArgumentException("Parameter cannot be null", "Owner");
                }

                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                var web = clientContext.Web;
                var siteUsers = web.SiteUsers;
                var associatedVisitorGroup = web.AssociatedVisitorGroup;
                var associatedMemberGroup = web.AssociatedMemberGroup;
                var associatedOwnerGroup = web.AssociatedOwnerGroup;

                clientContext.Load(siteUsers);
                clientContext.Load(associatedVisitorGroup, g => g.Title, g => g.Users);
                clientContext.Load(associatedMemberGroup, g => g.Title, g => g.Users);
                clientContext.Load(associatedOwnerGroup, g => g.Title, g => g.Users);
                clientContext.ExecuteQueryRetry();

                foreach(User user in associatedMemberGroup.Users)
                {
                    web.RemoveUserFromGroup(associatedMemberGroup, user);
                }

                foreach(User user in associatedOwnerGroup.Users)
                {
                    web.RemoveUserFromGroup(associatedOwnerGroup, user);
                }

                clientContext.ExecuteQueryRetry();

                web.AddUserToGroup(associatedOwnerGroup, request.Owner);
                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteReadOnlyResponse>(new SetSiteReadOnlyResponse { SetReadOnly = true }, new JsonMediaTypeFormatter())
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

        public class SetSiteReadOnlyRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
            [Required]
            [Display(Description = "Owner")]
            public string Owner { get; set; }
        }

        public class SetSiteReadOnlyResponse
        {
            [Display(Description = "True if the site was set to read-only")]
            public bool SetReadOnly { get; set; }
        }
    }
}

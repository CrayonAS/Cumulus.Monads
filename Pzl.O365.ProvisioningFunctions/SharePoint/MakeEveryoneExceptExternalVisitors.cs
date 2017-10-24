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
using Pzl.O365.ProvisioningFunctions.Helpers;

namespace Pzl.O365.ProvisioningFunctions.SharePoint
{
    public static class MakeEveryoneExceptExternalVisitors
    {
        [FunctionName("MakeEveryoneExceptExternalVisitors")]
        [ResponseType(typeof(MakeEveryoneExceptExternalVisitorsResponse))]
        [Display(Name = "Move Everyone (except external) users from member to visitor", Description = "In a public group make everyone visitors and not contributors.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]MakeEveryoneExceptExternalVisitorsRequest request, TraceWriter log)
        {
            string siteUrl = request.SiteURL;

            try
            {
                bool moved = await MoveEveryoneUser(log, siteUrl);

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<MakeEveryoneExceptExternalVisitorsResponse>(new MakeEveryoneExceptExternalVisitorsResponse { EveryOneExceptExternalMoved = moved }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        private static async Task<bool> MoveEveryoneUser(TraceWriter log, string siteUrl)
        {
            var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
            const string everyoneIdent = "c:0-.f|rolemanager|spo-grid-all-users/";
            bool moved = false;

            var web = clientContext.Web;
            var membersGroup = web.AssociatedMemberGroup;
            var siteUsers = web.SiteUsers;
            var visitorsGroup = web.AssociatedVisitorGroup;

            clientContext.Load(siteUsers);
            clientContext.Load(membersGroup);
            clientContext.Load(visitorsGroup);
            clientContext.ExecuteQueryRetry();
            foreach (User user in siteUsers)
            {
                if (!user.LoginName.StartsWith(everyoneIdent)) continue;

                if (web.IsUserInGroup(membersGroup.Title, user.LoginName))
                {
                    web.RemoveUserFromGroup(membersGroup, user);
                    web.AddUserToGroup(visitorsGroup, user);
                    moved = true;
                }
                break;
            }
            return moved;
        }

        public class MakeEveryoneExceptExternalVisitorsRequest
        {
            [Required]
            [Display(Description = "URL of site")]
            public string SiteURL { get; set; }
        }

        public class MakeEveryoneExceptExternalVisitorsResponse
        {
            [Display(Description = "Everyone group was moved from member to visitor")]
            public bool EveryOneExceptExternalMoved { get; set; }
        }
    }
}

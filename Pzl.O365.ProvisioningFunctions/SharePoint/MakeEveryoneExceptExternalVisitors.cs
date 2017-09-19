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
                var clientContext = await ConnectADAL.GetClientContext(siteUrl, log);
                string everyoneIdent = "c:0-.f|rolemanager|spo-grid-all-users/";
                bool moved = false;

                var web = clientContext.Web;
                clientContext.Load(web.SiteUsers);
                clientContext.ExecuteQueryRetry();
                foreach (User user in web.SiteUsers)
                {
                    if (user.LoginName.StartsWith(everyoneIdent))
                    {
                        var membersGroup = web.AssociatedMemberGroup;
                        clientContext.Load(membersGroup);
                        clientContext.ExecuteQueryRetry();

                        if (web.IsUserInGroup(membersGroup.Title, user.LoginName ))
                        {
                            web.RemoveUserFromGroup(membersGroup, user);
                            var visitorsGroup = web.AssociatedVisitorGroup;
                            web.AddUserToGroup(visitorsGroup, user);
                            moved = true;
                        }
                        break;
                    }
                }
                

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<MakeEveryoneExceptExternalVisitorsResponse>(new MakeEveryoneExceptExternalVisitorsResponse { EveryOneExceptExternalMoved = moved }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
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

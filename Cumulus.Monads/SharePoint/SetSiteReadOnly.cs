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
using System.Collections.Generic;

namespace Cumulus.Monads.SharePoint
{
    public static class SetSiteReadOnly
    {
        [FunctionName("SetSiteReadOnly")]
        [ResponseType(typeof(SetSiteReadOnlyResponse))]
        [Display(Name = "Set the site to read-only", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetSiteReadOnlyRequest request, TraceWriter log)
        {
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

                var clientContext = await ConnectADAL.GetClientContext(request.SiteURL, log);
                var web = clientContext.Web;
                const string everyoneIdent = "c:0-.f|rolemanager|spo-grid-all-users/";

                clientContext.Load(web,
                    w => w.AllProperties,
                    w => w.SiteUsers,
                    w => w.AssociatedVisitorGroup,
                    w => w.AssociatedMemberGroup,
                    w => w.AssociatedOwnerGroup);
                clientContext.ExecuteQueryRetry();

                var visitors = web.AssociatedVisitorGroup.Users;
                var members = web.AssociatedMemberGroup.Users;
                var owners = web.AssociatedOwnerGroup.Users;

                var visitorsPrivate = new List<User>();

                for (var i = 0; i < visitors.Count; i--)
                {
                    if (request.RemoveVisitors)
                    {
                        log.Info($"Removing {visitors[i].LoginName} from {web.AssociatedVisitorGroup.Title}");
                        web.RemoveUserFromGroup(web.AssociatedVisitorGroup, visitors[i]);
                    }
                }

                for (var i = 0; i < members.Count; i--)
                {
                    if (request.RemoveMembers)
                    {
                        log.Info($"Removing {members[i].LoginName} from {web.AssociatedMemberGroup.Title}");
                        web.RemoveUserFromGroup(web.AssociatedMemberGroup, members[i]);
                    }
                    visitorsPrivate.Add(members[i]);
                }

                for (var i = 0; i < owners.Count; i--)
                {
                    if (request.RemoveOwners)
                    {
                        log.Info($"Removing {owners[i].LoginName} from {web.AssociatedOwnerGroup.Title}");
                        web.RemoveUserFromGroup(web.AssociatedOwnerGroup, owners[i]);
                    }
                    visitorsPrivate.Add(owners[i]);
                }

                clientContext.ExecuteQueryRetry();

                log.Info($"Adding {request.Owner} to {web.AssociatedOwnerGroup.Title}");
                web.AddUserToGroup(web.AssociatedOwnerGroup, request.Owner);


                if (web.AllProperties.FieldValues.ContainsKey("GroupType") && web.AllProperties.FieldValues["GroupType"].ToString().Equals("Private"))
                {
                    log.Info($"The site is connected to a private group. Adding existing members/owners to {web.AssociatedVisitorGroup.Title}.");
                    for (var i = (visitorsPrivate.Count - 1); i >= 0; i--)
                    {
                        var user = visitorsPrivate[i];
                        if (user.LoginName.Contains("#ext#") && request.RemoveExternalUsers)
                        {
                            log.Info($"{user.LoginName} is an external user and will not be added to visitors.");
                        }
                        else
                        {
                            log.Info($"Adding {user.LoginName} to {web.AssociatedVisitorGroup.Title}.");
                            web.AddUserToGroup(web.AssociatedVisitorGroup, user);
                        }
                    }
                }
                else
                {
                    foreach (User user in web.SiteUsers)
                    {
                        if (user.LoginName.StartsWith(everyoneIdent))
                        {
                            log.Info($"Adding {user.LoginName} to {web.AssociatedVisitorGroup.Title}");
                            web.AddUserToGroup(web.AssociatedVisitorGroup, user);
                        }
                    }
                }

                clientContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<SetSiteReadOnlyResponse>(new SetSiteReadOnlyResponse { SetReadOnly = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Info(e.StackTrace);
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
            [Required]
            [Display(Description = "Remove external users")]
            public bool RemoveExternalUsers { get; set; }
            [Required]
            [Display(Description = "Remove users from members group")]
            public bool RemoveMembers { get; set; }
            [Required]
            [Display(Description = "Remove users from owners group")]
            public bool RemoveOwners { get; set; }
            [Required]
            [Display(Description = "Remove users from visitors group")]
            public bool RemoveVisitors { get; set; }
        }

        public class SetSiteReadOnlyResponse
        {
            [Display(Description = "True if the site was set to read-only")]
            public bool SetReadOnly { get; set; }
        }
    }
}

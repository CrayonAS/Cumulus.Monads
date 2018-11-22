using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Cumulus.Monads.Helpers;
using Microsoft.Graph;

namespace Cumulus.Monads.Graph
{
    public static class RemoveGroupMembers
    {
        [FunctionName("RemoveGroupMembers")]
        [ResponseType(typeof(RemoveGroupMembersResponse))]
        [Display(Name = "Remove group members", Description = "")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]RemoveGroupMembersRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.GroupId))
                {
                    throw new ArgumentException("Parameter cannot be null", "GroupId");
                }
                GraphServiceClient client = ConnectADAL.GetGraphClient(GraphEndpoint.v1);
                var group = client.Groups[request.GroupId];
                var members = await group.Members.Request().Select("displayName, id, mail, userPrincipalName, userType").GetAsync();
                var users = members.CurrentPage.Where(p => p.GetType() == typeof(User)).Cast<User>().ToList();

                // Removing users from group members
                for (int i = 0; i < users.Count; i++)
                {
                    var user = users[i];
                    log.Info($"Removing user {user.Id} from group {request.GroupId}");
                    await group.Members[user.Id].Reference.Request().DeleteAsync();
                }


                var removedGuestUsers = new List<User>();

                // Removes guest users
                for (int i = 0; i < users.Count; i++)
                {
                    var user = users[i];
                    log.Info($"Retrieving unified membership for user {user.Id}");
                    var memberOfPage = await client.Users[user.Id].MemberOf.Request().GetAsync();
                    var unifiedGroups = memberOfPage.CurrentPage.Where(p => p.GetType() == typeof(Group)).Cast<Group>().ToList().Where(g => g.GroupTypes.Contains("Unified")).ToList();
                    if(request.RemoveGuestUsers && user.UserType.Equals("Guest") && unifiedGroups.Count == 0)
                    {
                        log.Info($"Removing guest user {user.Id}");
                        await client.Users[user.Id].Request().DeleteAsync();
                        removedGuestUsers.Add(user);
                    }
                }
                var removeGroupMembersResponse = new RemoveGroupMembersResponse {
                    RemovedMembers = users,
                    RemovedGuestUsers = removedGuestUsers
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<RemoveGroupMembersResponse>(removeGroupMembersResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public static async Task RemoveUsersFromGroup(string groupId, IGroupRequestBuilder group, List<User> users, TraceWriter log)
        {
            for (int i = 0; i < users.Count; i++)
            {
                var user = users[i];
                log.Info($"Removing user {user.Id} from group {groupId}");
                await group.Members[user.Id].Reference.Request().DeleteAsync();
            }
        }

        public class RemoveGroupMembersRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
            [Display(Description = "Should guest users with no remaining Unified membership be removed from AD")]
            public bool RemoveGuestUsers { get; set; }
        }

        public class RemoveGroupMembersResponse
        {
            [Display(Description = "List of removed members")]
            public List<User> RemovedMembers { get; set; }
            [Display(Description = "List of removed guest users")]
            public List<User> RemovedGuestUsers { get; set; }
        }
    }
}

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
                GraphServiceClient client = ConnectADAL.GetGraphClient();
                var group = client.Groups[request.GroupId];
                var members = await group.Members.Request().GetAsync();
                for (int i = 0; i < members.Count; i++)
                {
                    var member = members[i];
                    log.Info($"Removing user {member.Id} from group {request.GroupId}");
                    await group.Members[member.Id].Reference.Request().DeleteAsync();
                }
                var removeGroupMembersResponse = new RemoveGroupMembersResponse { RemovedMembers = members };
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

        public class RemoveGroupMembersRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class RemoveGroupMembersResponse
        {
            [Display(Description = "True/false if members was removed")]
            public IGroupMembersCollectionWithReferencesPage RemovedMembers { get; set; }
        }
    }
}

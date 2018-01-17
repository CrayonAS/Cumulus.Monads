using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Web.Http.Description;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Pzl.O365.ProvisioningFunctions.Helpers;
using Group = Microsoft.Graph.Group;

namespace Pzl.O365.ProvisioningFunctions.Graph
{
    public static class CreateGroup
    {

        [FunctionName("CreateGroup")]
        [ResponseType(typeof(CreateGroupResponse))]
        [Display(Name = "Create Office 365 Group", Description = "This action will create a new Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]CreateGroupRequest request, TraceWriter log)
        {
            try
            {
                string mailNickName = await GetUniqueMailAlias(request.Name, request.Prefix);
                GraphServiceClient client = ConnectADAL.GetGraphClient();
                var newGroup = new Group
                {
                    DisplayName = GetDisplayName(request.Name, request.Prefix),
                    Description = "",
                    MailNickname = mailNickName,
                    MailEnabled = true,
                    SecurityEnabled = false,
                    Visibility = request.Public ? "Public" : "Private",
                    GroupTypes = new List<string> { "Unified" },
                };

                var addedGroup = await client.Groups.Request().AddAsync(newGroup);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<CreateGroupResponse>(new CreateGroupResponse{ GroupId = addedGroup.Id }, new JsonMediaTypeFormatter())
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

        static string GetDisplayName(string name, string prefix)
        {
             string displayName = Regex.Replace(name, @":?\s+", "", RegexOptions.IgnoreCase);
             if(string.IsNullOrWhiteSpace(prefix)) {
                return displayName;
             } else {
                return $"{prefix}: {displayName}";
             }
        }

        static async Task<string> GetUniqueMailAlias(string name, string prefix = "")
        {
            var mailNickname = Regex.Replace(name.ToLower(), @":?\s+", "", RegexOptions.IgnoreCase);
            mailNickname = Regex.Replace(mailNickname, "[^a-z0-9]", "");
            if(string.IsNullOrWhiteSpace(prefix)) {
                mailNickname = mailNickname.ToLower();
            } else {
                mailNickname = $"{prefix}-{mailNickname}".ToLower();
            }
            if (string.IsNullOrWhiteSpace(mailNickname))
            {
                mailNickname = new Random().Next(0, 9).ToString();
            }
            const int maxCharsInEmail = 40;
            if (mailNickname.Length > maxCharsInEmail)
            {
                mailNickname = mailNickname.Substring(0, maxCharsInEmail);
            }

            GraphServiceClient client = ConnectADAL.GetGraphClient();
            while (true)
            {
                IGraphServiceGroupsCollectionPage groupExist = await client.Groups.Request()
                    .Filter($"groupTypes/any(grp: grp eq 'Unified') and MailNickname eq '{mailNickname}'").Top(1)
                    .GetAsync();
                if (groupExist.Count > 0)
                {
                    mailNickname += new Random().Next(0, 9).ToString();
                }
                else
                {
                    break;
                }
            }
            return mailNickname;
        }

        public class CreateGroupRequest
        {
            [Required]
            public string Name { get; set; }
            public string Description { get; set; }
            public string Prefix { get; set; }
            [Required]
            public string Responsible { get; set; }
            [Required]
            public bool Public { get; set; }
        }

        public class CreateGroupResponse
        {
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }
    }
}

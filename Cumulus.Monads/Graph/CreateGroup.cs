using System;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Web.Http.Description;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Graph;
using Cumulus.Monads.Helpers;
using Group = Microsoft.Graph.Group;

namespace Cumulus.Monads.Graph
{
    public static class CreateGroup
    {
        private static readonly Regex ReRemoveIllegalChars = new Regex("[^a-z0-9-.]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        [FunctionName("CreateGroup")]
        [ResponseType(typeof(CreateGroupResponse))]
        [Display(Name = "Create Office 365 Group", Description = "This action will create a new Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]CreateGroupRequest request, TraceWriter log)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.Name))
                {
                    throw new ArgumentException("Parameter cannot be null", "Name");
                }
                if (string.IsNullOrWhiteSpace(request.Description))
                {
                    throw new ArgumentException("Parameter cannot be null", "Description");
                }
                string mailNickName = await GetUniqueMailAlias(request);
                string displayName = GetDisplayName(request);
                GraphServiceClient client = ConnectADAL.GetGraphClient(GraphEndpoint.Beta);
                var newGroup = new Group
                {
                    DisplayName = displayName,
                    Description = GetDescription(request.Description, 1000),
                    MailNickname = mailNickName,
                    MailEnabled = true,
                    SecurityEnabled = false,
                    Visibility = request.Public ? "Public" : "Private",
                    GroupTypes = new List<string> { "Unified" },
                    Classification = request.Classification
                };
                var addedGroup = await client.Groups.Request().AddAsync(newGroup);
                var createGroupResponse = new CreateGroupResponse
                {
                    GroupId = addedGroup.Id,
                    DisplayName = displayName,
                    Mail = addedGroup.Mail
                };
                if (request.AllowToAddGuests != null)
                {
                    var groupUnifiedGuestSetting = new GroupSetting()
                    {
                        DisplayName = "Group.Unified.Guest",
                        TemplateId = "08d542b9-071f-4e16-94b0-74abb372e3d9",
                        Values = new List<SettingValue> { new SettingValue() { Name = "AllowToAddGuests", Value = request.AllowToAddGuests.ToString() } }
                    };
                    log.Info($"Setting setting in Group.Unified.Guest (08d542b9-071f-4e16-94b0-74abb372e3d9), AllowToAddGuests = {request.AllowToAddGuests.ToString()}");
                    await client.Groups[addedGroup.Id].Settings.Request().AddAsync(groupUnifiedGuestSetting);
                }
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<CreateGroupResponse>(createGroupResponse, new JsonMediaTypeFormatter())
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

        static string GetDisplayName(CreateGroupRequest request)
        {
            string prefix = string.Empty;
            string suffix = string.Empty;
            var displayName = request.Name;
            var prefixSeparator = string.Empty;
            CultureInfo cultureInfo = Thread.CurrentThread.CurrentCulture;

            if (!string.IsNullOrWhiteSpace(request.Prefix) && request.UsePrefixInDisplayName)
            {
                //remove prefix from name if accidentally added as part of the name
                displayName = Regex.Replace(displayName, "^" + request.Prefix + @":?\s+", "", RegexOptions.IgnoreCase);
                prefix = cultureInfo.TextInfo.ToTitleCase(request.Prefix);
                prefixSeparator = ":";
            }

            if (!string.IsNullOrWhiteSpace(request.Suffix) && request.UseSuffixInDisplayName)
            {
                suffix = cultureInfo.TextInfo.ToTitleCase(request.Suffix);                
            }
            displayName = $"{prefix}{prefixSeparator} {displayName} {suffix}".Trim();
            return displayName;
        }

        static string GetDescription(string description, int maxLength)
        {
            return description.Length > maxLength ? description.Substring(0, maxLength) : description;
        }

        static async Task<string> GetUniqueMailAlias(CreateGroupRequest request)
        {
            string name = string.IsNullOrEmpty(request.Alias) ? request.Name : request.Alias;
            string prefix = request.Prefix;
            string suffix = request.Suffix;
            string mailNickname = ReRemoveIllegalChars.Replace(name, "").ToLower();
            prefix = ReRemoveIllegalChars.Replace(prefix + "", "").ToLower();
            suffix = ReRemoveIllegalChars.Replace(suffix + "", "").ToLower();

            string prefixSeparator = string.Empty;
            if (!string.IsNullOrWhiteSpace(prefix) && request.UsePrefixInMailAlias)
            {
                prefixSeparator = string.IsNullOrWhiteSpace(request.PrefixSeparator) ? "-" : request.PrefixSeparator;                
            }
            string suffixSeparator = string.Empty;
            if (!string.IsNullOrWhiteSpace(suffix) && request.UseSuffixInMailAlias)
            {
                suffixSeparator = string.IsNullOrWhiteSpace(request.SuffixSeparator) ? "-" : request.SuffixSeparator;               
            }

            int maxCharsInEmail = 40 - prefix.Length - prefixSeparator.Length - suffixSeparator.Length - suffix.Length;
            if (mailNickname.Length > maxCharsInEmail)
            {
                mailNickname = mailNickname.Substring(0, maxCharsInEmail);
            }

            mailNickname = $"{prefix}{prefixSeparator}{mailNickname}{suffixSeparator}{suffix}";

            if (string.IsNullOrWhiteSpace(mailNickname))
            {
                mailNickname = new Random().Next(0, 9).ToString();
            }

            GraphServiceClient client = ConnectADAL.GetGraphClient();
            while (true)
            {
                IGraphServiceGroupsCollectionPage groupExist = await client.Groups.Request()
                    .Filter($"groupTypes/any(grp: grp eq 'Unified') and MailNickname eq '{mailNickname}'").Top(1)
                    .GetAsync();
                if (groupExist.Count > 0)
                {
                    string number = new Random().Next(0, 9).ToString();                    
                    if (string.IsNullOrWhiteSpace(suffixSeparator + suffix))
                    {
                        mailNickname += new Random().Next(0, 9).ToString();
                    }
                    else
                    {
                        int suffixIdx = mailNickname.IndexOf(suffixSeparator + suffix);
                        mailNickname = mailNickname.Insert(suffixIdx, number);
                    }
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
            [Display(Description = "Name of the group")]
            public string Name { get; set; }

            [Display(Description = "E-mail alias for the group")]
            public string Alias { get; set; }

            [Required]
            [Display(Description = "Description of the group")]
            public string Description { get; set; }

            [Display(Description = "Prefix for group display name / e-mail address")]
            public string Prefix { get; set; }

            [Display(Description = "Separator character between prefix and name")]
            public string PrefixSeparator { get; set; }

            [Display(Description = "Suffix for group display name / e-mail address")]
            public string Suffix { get; set; }

            [Display(Description = "Separator character between suffix and name")]
            public string SuffixSeparator { get; set; }

            [Required]
            [Display(Description = "Should the group be public")]
            public bool Public { get; set; }

            [Display(Description = "If prefix is set, use for DisplayName")]
            public bool UsePrefixInDisplayName { get; set; }

            [Display(Description = "If prefix is set, use for EmailAlias")]
            public bool UsePrefixInMailAlias { get; set; }

            [Display(Description = "If suffix is set, use for EmailAlias")]
            public bool UseSuffixInMailAlias { get; set; }

            [Display(Description = "If suffix is set, use for DisplayName")]
            public bool UseSuffixInDisplayName { get; set; }

            [Display(Description = "Classification")]
            public string Classification { get; set; }

            [Display(Description = "Classification")]
            public bool? AllowToAddGuests { get; set; }
        }

        public class CreateGroupResponse
        {
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Display(Description = "DisplayName of the Office 365 Group")]
            public string DisplayName { get; set; }

            [Display(Description = "Mail of the Office 365 Group")]
            public string Mail { get; set; }
        }
    }
}

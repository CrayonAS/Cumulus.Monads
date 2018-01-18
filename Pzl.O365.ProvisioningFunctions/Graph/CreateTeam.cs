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
    public static class CreateTeam
    {

        [FunctionName("CreateTeam")]
        [ResponseType(typeof(CreateTeamResponse))]
        [Display(Name = "Create Office 365 Team", Description = "This action will create a new Team for the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]CreateTeamRequest request, TraceWriter log)
        {
            try
            {               
                var createTeamResponse = new CreateTeamResponse
                {
                    TeamCreated = true
                };
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<CreateTeamResponse>(createTeamResponse, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error:  {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ExpectationFailed)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class CreateTeamRequest
        {
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class CreateTeamResponse
        {
            public bool TeamCreated { get; set; }
        }
    }
}

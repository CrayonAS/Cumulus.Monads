using System;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Net.Http;
using System.Net.Http.Formatting;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http.Description;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using Pzl.O365.ProvisioningFunctions.Helpers;

namespace Pzl.O365.ProvisioningFunctions.Graph
{
    //TODO: Change to proper GraphClient support once classification moves from Beta endpoint
    public static class ApplyLifecyclePolicy
    {
        [FunctionName("ApplyLifecyclePolicy")]
        [ResponseType(typeof(ApplyLifecyclePolicyResponse))]
        [Display(Name = "Apply lifecycle policy to an Office 365 Group", Description = "Apply an expiration lifecyle policy to the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]ApplyLifecyclePolicyRequest request, TraceWriter log)
        {
            try
            {
                string bearerToken = await ConnectADAL.GetBearerTokenServiceIdentity(log);
                log.Info(bearerToken);
                dynamic template = new { groupId = request.GroupId };
                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/beta/groupLifecyclePolicies/aa31e487-96a4-4a23-ae25-a5ba3e51e815/addGroup");
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PostAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<ApplyLifecyclePolicyResponse>(new ApplyLifecyclePolicyResponse { IsApplied = true }, new JsonMediaTypeFormatter())
                    });
                }

                string responseMsg = await response.Content.ReadAsStringAsync();
                dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<object>(errorMsg, new JsonMediaTypeFormatter())
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

        public class ApplyLifecyclePolicyRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
        }

        public class ApplyLifecyclePolicyResponse
        {
            [Display(Description = "true/false if applied")]
            public bool IsApplied { get; set; }
        }
    }
}

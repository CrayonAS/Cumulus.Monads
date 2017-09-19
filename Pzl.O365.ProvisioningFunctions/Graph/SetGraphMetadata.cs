using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Dynamic;
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
    public static class SetGraphMetadata
    {
        [FunctionName("SetGraphMetadata")]
        [ResponseType(typeof(SetGraphMetadataResponse))]
        [Display(Name = "Set Office 365 Group metadata", Description = "Store metadata for the Office 365 Group in the Microsoft Graph")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGraphMetadataRequest request, TraceWriter log)
        {
            try
            {
                string extensionName = await SchemaExtensionHelper.GetExtensionName(ConnectADAL.AppId);
                if (string.IsNullOrWhiteSpace(extensionName))
                {
                    throw new Exception($"{extensionName} not found");
                }

                dynamic property = new ExpandoObject();
                ((IDictionary<string, object>)property).Add(request.Key.ToString(), request.Value);
                dynamic template = new ExpandoObject();
                ((IDictionary<string, object>)template).Add(extensionName, property);

                var content = new StringContent(JsonConvert.SerializeObject(template), Encoding.UTF8, "application/json");
                Uri uri = new Uri($"https://graph.microsoft.com/v1.0/groups/{request.GroupId}");
                string bearerToken = await ConnectADAL.GetBearerToken();
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                var response = await client.PatchAsync(uri, content);
                if (response.IsSuccessStatusCode)
                {
                    return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ObjectContent<SetGraphMetadataResponse>(new SetGraphMetadataResponse { Added = true }, new JsonMediaTypeFormatter())
                    });
                }
                string responseMsg = await response.Content.ReadAsStringAsync();
                dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new ObjectContent<object>(errorMsg, new JsonMediaTypeFormatter())
                });

            }
            catch (Exception e)
            {
                log.Error(e.Message);
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }


        public enum MetadataFields
        {
            GroupType = 0,
            Responsible = 1,
            StartDate = 2,
            EndDate = 3
        }

        public class SetGraphMetadataRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }

            [Required]
            [Display(Description = "Metadata name. Valid values are: groupType / responsible")]
            public MetadataFields Key { get; set; }

            [Required]
            [Display(Description = "Metadata value")]
            public string Value { get; set; }
        }
    }

    public class SetGraphMetadataResponse
    {
        [Display(Description = "true/false if added")]
        public bool Added { get; set; }
    }
}

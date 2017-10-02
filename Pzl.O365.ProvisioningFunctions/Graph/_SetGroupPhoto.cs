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
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Pzl.O365.ProvisioningFunctions.Helpers;

namespace Pzl.O365.ProvisioningFunctions.Graph
{
    //TODO: Comment in when service principal works against graph
    public static class SetGroupPhoto
    {
        //[FunctionName("SetGroupPhoto")]
        [ResponseType(typeof(SetGroupPhotoResponse))]
        [Display(Name = "Set logo for an Office 365 Group", Description = "Set the logo image for the Office 365 Group")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]SetGroupPhotoRequest request, TraceWriter log)
        {
            try
            {
                GetGroupSite.GetGroupSiteRequest getGroup = new GetGroupSite.GetGroupSiteRequest { GroupId = request.GroupId };
                var groupSite = await GetGroupSite.Run(getGroup, log);
                log.Info("Got Group URL");

                var clientContext = await ConnectADAL.GetClientContext(groupSite.SiteURL, log);

                Uri fileUri = new Uri(request.LogoURL);
                var webUrl = Web.WebUrlFromFolderUrlDirect(clientContext, fileUri);
                var logoContext = clientContext.Clone(webUrl.ToString());
                var file = logoContext.Web.GetFileByUrl(request.LogoURL);
                logoContext.Load(file);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                logoContext.ExecuteQueryRetry();

                log.Info("Got logo image");

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        mStream.Position = 0;
                        Uri uri = new Uri($"https://graph.microsoft.com/beta/groups/{request.GroupId}/photo/$value");

                        ConnectADAL.MsiInformation serviceInfo = await ConnectADAL.GetBearerTokenServiceIdentity(log);

                        HttpClient client = new HttpClient();
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", serviceInfo.BearerToken);
                        var content = new StreamContent(mStream);
                        content.Headers.ContentType = MediaTypeHeaderValue.Parse("image/jpeg");
                        var response = await client.PutAsync(uri, content);

                        if (response.IsSuccessStatusCode)
                        {
                            return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                            {
                                Content = new ObjectContent<SetGroupPhotoResponse>(new SetGroupPhotoResponse { IsUpdated = true }, new JsonMediaTypeFormatter())
                            });
                        }
                        string responseMsg = await response.Content.ReadAsStringAsync();
                        dynamic errorMsg = JsonConvert.DeserializeObject(responseMsg);
                        return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                        {
                            Content = new ObjectContent<object>(errorMsg, new JsonMediaTypeFormatter())
                        });
                    }
                }
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new ObjectContent<string>("Image not found", new JsonMediaTypeFormatter())
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

        public class SetGroupPhotoRequest
        {
            [Required]
            [Display(Description = "Id of the Office 365 Group")]
            public string GroupId { get; set; }
            [Required]
            [Display(Description = "SharePoint URL to logo")]
            public string LogoURL { get; set; }
        }

        public class SetGroupPhotoResponse
        {
            [Display(Description = "true/false if set")]
            public bool IsUpdated { get; set; }
        }
    }
}

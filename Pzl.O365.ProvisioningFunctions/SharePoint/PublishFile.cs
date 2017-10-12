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
    public static class PublishFile
    {
        [FunctionName("PublishFile")]
        [ResponseType(typeof(PublishFileResponse))]
        [Display(Name = "Publish a file SharePoint file", Description = "Make sure a file is published as a major version.")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "post")]PublishFileRequest request, TraceWriter log)
        {
            string fileName = System.IO.Path.GetFileName(request.FileURL);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                log.Error($"Error: filename is missing");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>("Filename is missing", new JsonMediaTypeFormatter())
                });
            }
            Uri fileUri = new Uri(request.FileURL.Replace(fileName, ""));
            string rootUrl = $"{fileUri.Scheme}://{fileUri.Authority}";

            var clientContext = await ConnectADAL.GetClientContext(rootUrl, log);
            var webUrl = Web.WebUrlFromFolderUrlDirect(clientContext, fileUri);
            var fileContext = clientContext.Clone(webUrl.ToString());

            try
            {
                var file = fileContext.Web.GetFileByUrl(request.FileURL);
                fileContext.Load(file);
                fileContext.ExecuteQueryRetry();
                if (file.CheckOutType != CheckOutType.None)
                {
                    file.UndoCheckOut();
                }

                file.CheckOut();
                file.CheckIn("Updated", CheckinType.MajorCheckIn);
                fileContext.ExecuteQueryRetry();

                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ObjectContent<PublishFileResponse>(new PublishFileResponse { Published = true }, new JsonMediaTypeFormatter())
                });
            }
            catch (Exception e)
            {
                log.Error($"Error: {e.Message }\n\n{e.StackTrace}");
                return await Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
                {
                    Content = new ObjectContent<string>(e.Message, new JsonMediaTypeFormatter())
                });
            }
        }

        public class PublishFileRequest
        {
            [Required]
            [Display(Description = "URL of file")]
            public string FileURL { get; set; }
        }

        public class PublishFileResponse
        {
            [Display(Description = "True if file was published")]
            public bool Published { get; set; }
        }
    }
}

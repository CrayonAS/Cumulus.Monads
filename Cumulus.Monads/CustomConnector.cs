using System.Dynamic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;

namespace Cumulus.Monads
{
    public static class CustomConnector
    {
        [FunctionName("CustomConnector")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {


            var assembly = Assembly.GetExecutingAssembly();

            dynamic swaggerobj = new ExpandoObject();
            swaggerobj.swagger = "2.0";
            swaggerobj.info = new ExpandoObject();
            swaggerobj.info.title = assembly.GetName().Name;
            swaggerobj.info.version = "1.0.0";
            swaggerobj.host = req.RequestUri.Authority;
            swaggerobj.basePath = "/";
            swaggerobj.schemes = new[] { "https" };
            if (swaggerobj.host.Contains("127.0.0.1") || swaggerobj.host.Contains("localhost"))
            {
                swaggerobj.schemes = new[] { "http" };
            }
            swaggerobj.definitions = new ExpandoObject();
            swaggerobj.paths = Cumulus.Monads.Swagger.GeneratePaths(assembly, swaggerobj);
            swaggerobj.securityDefinitions = Cumulus.Monads.Swagger.GenerateSecurityDefinitions();
           
            var jsonobj = JsonConvert.SerializeObject(swaggerobj);
          
            var swaggerString = $@"The Swagger File Contents will Appear Hear";
            var template = @"{'$schema': 'https://schema.management.azure.com/schemas/2015-01-01/deploymentTemplate.json#', 'contentVersion': '1.0.0.0', 
               'parameters': {}, 'variables': {}, 
               'resources': [
                    {
                      'apiVersion': '2016-06-01',
                      'dependsOn': [],
                      'location': '[resourceGroup().location]',
                      'name': 'CumulusConnectorTest1',
                      'properties': {
                        'connectionParameters': {
                          'api_key': {
                            'type': 'securestring',
                            'uiDefinition': {
                              'displayName': 'Function Host Key',
                              'description': 'The Function App Key for this api',
                              'tooltip': 'Provide your Function App Key',
                              'constraints': {
                                'tabIndex': 2,
                                'clearText': false,
                                'required': 'true'
                              }
                            }
                          }
                        },
                        'swagger' :  " + jsonobj + @",
                        'displayName': 'Cumulus Connector Test',
                        
                        'backendService': {
                          'serviceUrl': 'https://cumulusmonadssecond.azurewebsites.net'
                        },
                        'apiType': 'Rest'
                      },
                      'type': 'Microsoft.Web/customApis'
                    }
                ],
               'outputs': {}
            }";
        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
        response.Content = new StringContent(template, System.Text.Encoding.UTF8, "application/json");
        return response;
        }
    }
}

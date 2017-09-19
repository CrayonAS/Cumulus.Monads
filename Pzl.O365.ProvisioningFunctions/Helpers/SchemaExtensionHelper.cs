using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph;

namespace Pzl.O365.ProvisioningFunctions.Helpers
{
    class SchemaExtensionHelper
    {
        private const string MetadataSchemaName = "CollaborationMetadata";
        public static async Task<string> GetExtensionName(string ownerId, string schemaName = MetadataSchemaName)
        {
            GraphServiceClient graphClient = ConnectADAL.GetGraphClient();
            IGraphServiceSchemaExtensionsCollectionPage tenantSchemas = await graphClient.SchemaExtensions.Request().Filter($"owner eq '{ownerId}'").GetAsync();
            string extensionName = null;
            bool moreData = true;
            while (moreData)
            {
                foreach (SchemaExtension schemaExtension in tenantSchemas)
                {
                    if (!schemaExtension.Id.EndsWith(schemaName) || schemaExtension.Status != "Available") continue;
                    extensionName = schemaExtension.Id;
                    break;
                }
                if (tenantSchemas.NextPageRequest != null)
                {
                    tenantSchemas = await tenantSchemas.NextPageRequest.GetAsync();
                }
                else
                {
                    moreData = false;
                }
            }
            return extensionName;
        }
    }
}

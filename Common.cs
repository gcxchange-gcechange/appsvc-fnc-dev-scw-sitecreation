using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using System;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class Common
    {
       //public class SpaceRequest
       // {
       //     public string Id { get; set; }
       //     public string SecurityCategory { get; set; }
       //     public string SpaceName { get; set; }
       //     public string SpaceNameFR { get; set; }
       //     public string Owner1 { get; set; }
       //     public string SpaceDescription { get; set; }
       //     public string SpaceDescriptionFR { get; set; }
       //     public string TemplateTitle { get; set; }
       //     public string TeamPurpose { get; set; }
       //     public string BusinessJustification { get; set; }
       //     public string RequesterName { get; set; }
       //     public string RequesterEmail { get; set; }
       //     public string Status { get; set; }
       //     public string ApprovedDate { get; set; }
       //     public string Comment { get; set; }
       // }

        public static async Task InsertMessageAsync(string connectionString, string queueName, ListItem listItem, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            try {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference(queueName);

                string serializedMessage = JsonConvert.SerializeObject(listItem.Fields.AdditionalData);

                CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
                await queue.AddMessageAsync(message);
            }
            catch (Exception e) {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
            }

            log.LogInformation("InsertMessageAsync processed a request.");
        }
    }
}
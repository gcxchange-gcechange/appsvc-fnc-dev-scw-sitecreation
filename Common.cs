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
        public static async Task InsertMessageAsync(string connectionString, string queueName, ListItem listItem, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            try {
                CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                CloudQueue queue = queueClient.GetQueueReference(queueName);

                string serializedMessage = JsonConvert.SerializeObject(listItem.Fields.AdditionalData);

                log.LogInformation($"serializedMessage = {serializedMessage}");

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
using Azure.Storage.Queues;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Newtonsoft.Json;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class Common
    {
        public static async Task InsertMessageAsync(string connectionString, string queueName, ListItem listItem, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            try {
                string serializedMessage = JsonConvert.SerializeObject(listItem.Fields.AdditionalData);
                log.LogInformation($"serializedMessage = {serializedMessage}");

                QueueClientOptions options = new QueueClientOptions() { MessageEncoding = QueueMessageEncoding.Base64 };
                QueueClient client = new QueueClient(connectionString, queueName, options);

                await client.SendMessageAsync(serializedMessage);
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
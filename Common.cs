using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class Common
    {


        public class SpaceRequest
        {
            public string Id { get; set; }
            public string SpaceName { get; set; }
            public string SpaceNameFR { get; set; }
            public string Owner1 { get; set; }
            public string SpaceDescription { get; set; }
            public string SpaceDescriptionFR { get; set; }
            public string TemplateTitle { get; set; }
            public string TeamPurpose { get; set; }
            public string BusinessJustification { get; set; }
            public string RequesterName { get; set; }
            public string RequesterEmail { get; set; }
            public string Status { get; set; }
            public string ApprovedDate { get; set; }
            public string Comment { get; set; }
        }

        public static async Task InsertMessageAsync(string connectionString, string queueName, ListItem listItem, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference(queueName);

            SpaceRequest request = new SpaceRequest();

            request.Id = listItem.Id;

            //request.SpaceName = listItem.Fields.AdditionalData["Title"].ToString();
            //request.SpaceNameFR = listItem.Fields.AdditionalData["SpaceNameFR"].ToString();
            //request.Owner1 = listItem.Fields.AdditionalData["Owner1"].ToString();
            //request.SpaceDescription = listItem.Fields.AdditionalData["SpaceDescription"].ToString();
            //request.SpaceDescriptionFR = listItem.Fields.AdditionalData["SpaceDescriptionFR"].ToString();
            //request.TemplateTitle = listItem.Fields.AdditionalData["TemplateTitle"].ToString();
            //request.TeamPurpose = listItem.Fields.AdditionalData["TeamPurpose"].ToString();
            //request.BusinessJustification = listItem.Fields.AdditionalData["BusinessJustification"].ToString();
            //request.RequesterName = listItem.Fields.AdditionalData["RequesterName"].ToString();
            //request.RequesterEmail = listItem.Fields.AdditionalData["RequesterEmail"].ToString();
            //request.Status = listItem.Fields.AdditionalData["Status"].ToString();

            request.SpaceName = listItem.Fields.AdditionalData.Keys.Contains("Title") ? listItem.Fields.AdditionalData["Title"].ToString() : string.Empty;
            request.SpaceNameFR = listItem.Fields.AdditionalData.Keys.Contains("SpaceNameFR") ? listItem.Fields.AdditionalData["SpaceNameFR"].ToString() : string.Empty;
            request.Owner1 = listItem.Fields.AdditionalData.Keys.Contains("Owner1") ? listItem.Fields.AdditionalData["Owner1"].ToString() : string.Empty;
            request.SpaceDescription = listItem.Fields.AdditionalData.Keys.Contains("SpaceDescription") ? listItem.Fields.AdditionalData["SpaceDescription"].ToString() : string.Empty;
            request.SpaceDescriptionFR = listItem.Fields.AdditionalData.Keys.Contains("SpaceDescriptionFR") ? listItem.Fields.AdditionalData["SpaceDescriptionFR"].ToString() : string.Empty;
            request.TemplateTitle = listItem.Fields.AdditionalData.Keys.Contains("TemplateTitle") ? listItem.Fields.AdditionalData["TemplateTitle"].ToString() : string.Empty;
            request.TeamPurpose = listItem.Fields.AdditionalData.Keys.Contains("TeamPurpose") ? listItem.Fields.AdditionalData["TeamPurpose"].ToString() : string.Empty;
            request.BusinessJustification = listItem.Fields.AdditionalData.Keys.Contains("BusinessJustification") ? listItem.Fields.AdditionalData["BusinessJustification"].ToString() : string.Empty;
            request.RequesterName = listItem.Fields.AdditionalData.Keys.Contains("RequesterName") ? listItem.Fields.AdditionalData["RequesterName"].ToString() : string.Empty;
            request.RequesterEmail = listItem.Fields.AdditionalData.Keys.Contains("RequesterEmail") ? listItem.Fields.AdditionalData["RequesterEmail"].ToString() : string.Empty;
            request.Status = listItem.Fields.AdditionalData.Keys.Contains("Status") ? listItem.Fields.AdditionalData["Status"].ToString() : string.Empty;
            request.ApprovedDate = listItem.Fields.AdditionalData.Keys.Contains("ApprovedDate") ? listItem.Fields.AdditionalData["ApprovedDate"].ToString() : string.Empty;
            request.Comment = listItem.Fields.AdditionalData.Keys.Contains("Comment") ? listItem.Fields.AdditionalData["Comment"].ToString() : string.Empty;

            string serializedMessage = JsonConvert.SerializeObject(request);

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await queue.AddMessageAsync(message);

            log.LogInformation("InsertMessageAsync processed a request.");
        }



    }
}

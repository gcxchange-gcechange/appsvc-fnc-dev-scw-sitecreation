//using System;
//using System.IO;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.Azure.WebJobs;
//using Microsoft.Azure.WebJobs.Extensions.Http;
//using Microsoft.AspNetCore.Http;
//using Microsoft.Extensions.Logging;
//using Newtonsoft.Json;
//using Microsoft.Graph;
//using System.Collections.Generic;

//namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
//{
//    public static class SetUnclassified
//    {
//        [FunctionName("SetUnclassified")]
//        public static async Task<IActionResult> Run(
//            [HttpTrigger(AuthorizationLevel.System, "get", "post", Route = null)] HttpRequest req, ILogger log)
//        {
//            log.LogInformation("SetUnclassified trigger function processed a request.");

//            // Anon
//            // Function


//            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
//            dynamic data = JsonConvert.DeserializeObject(requestBody);


//            //// var scopes = new[] { "AuditLog.Read.All", "Directory.ReadWrite.All" };

//            Auth auth = new Auth();
//            log.LogInformation("------1------");
//            var graphAPIAuth = auth.graphAuth(log);
//            log.LogInformation("------2------");

//            var groupId = "9d907192-71d6-488e-addf-c566ff4ee2e7";
//            var labelId = "a1ab9d1a-185f-40cc-97d9-e1177019a70b";

//            log.LogInformation("------3------");

//            await AddUnclassifiedLabel(graphAPIAuth, labelId, groupId, log);

//            log.LogInformation("------4------");

//            return new OkResult();
//        }

//        public static async Task<string> AddUnclassifiedLabel(GraphServiceClient graphClient, string labelid, string groupid, ILogger log)
//        {
//            var group = new Group
//            {
//                AssignedLabels = new List<AssignedLabel>()
//                    {
//                        new AssignedLabel
//                        {
//                            LabelId = labelid
//                        }
//                    },
//            };

//            // groupid

//            log.LogInformation($"groupid: {groupid}");


//            try {
//                var users = await graphClient.Groups[groupid].Request().UpdateAsync(group);
//            }
//            catch (Exception e)
//            {
//                log.LogInformation($"Message: {e.Message}");
//                if (e.InnerException is not null)
//                    log.LogInformation($"InnerException: {e.InnerException.Message}");
//                return string.Empty;
//            }

//            //log.LogInformation($"users: {users}");
//            return "true";
//        }

//    }
//}
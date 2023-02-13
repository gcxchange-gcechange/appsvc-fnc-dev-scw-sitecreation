//using System;
//using System.Threading.Tasks;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.Azure.WebJobs;
//using Microsoft.Azure.WebJobs.Extensions.Http;
//using Microsoft.AspNetCore.Http;
//using Microsoft.Extensions.Logging;
//using Microsoft.Graph;
//using System.Collections.Generic;

//namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
//{
//    public static class AddTeamMembersTest
//    {
//        [FunctionName("AddTeamMembersTest")]
//        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
//        {
//            log.LogInformation("C# HTTP trigger function processed a request.");

//            Auth auth = new Auth();
//            var graphClient = auth.graphAuth(log);

//            await AddMembersToGroup(graphClient, log, "25a36164-466e-4ff1-aeb6-62fe760df92e", "44ec56a1-6066-49f2-a227-ec56eda7a7cd,331d028f-a1d5-4d66-9438-9aac013c36b9");
//            // 

//            return new OkResult();
//        }

//        public static async Task<bool> AddMembersToGroup(GraphServiceClient graphClient, ILogger log, string groupId, string Members)
//        {
//            try
//            {
//                //List<AadUserConversationMember> mems = new List<AadUserConversationMember>();

//                foreach (string memberId in Members.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
//                {
//                    AadUserConversationMember mem = new AadUserConversationMember
//                    {
//                        Roles = new List<String>()
//                        {
//                            "member"
//                        },
//                        AdditionalData = new Dictionary<string, object>()
//                        {
//                            {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{memberId}')"}
//                        }
//                    };

//                    log.LogInformation("Get ready!");
//                    await graphClient.Groups[groupId].Team.Members
//                        .Request()
//                        .AddAsync(mem);
//                    log.LogInformation("Good job!");
//                }
//            }
//            catch (Exception e)
//            {
//                log.LogError($"Message: {e.Message}");
//                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
//                log.LogError($"StackTrace: {e.StackTrace}");
//            }

//            log.LogInformation($"Licensed add to owner of {groupId} successfully.");

//            return true;
//        }
//    }
//}

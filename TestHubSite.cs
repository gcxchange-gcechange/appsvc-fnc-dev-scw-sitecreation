using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
//using System.Net.Http;
//using Microsoft.Identity.Client;
//using Microsoft.Graph;
//using static appsvc_fnc_dev_scw_sitecreation_dotnet001.Auth;
//using Microsoft.Extensions.Configuration;
//using System;
//using Azure.Core;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    public static class TestHubSite
    {
        [FunctionName("TestHubSite")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            //IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            //string delegatedUserName = config["delegatedUserName"];
            //string delegatedUserSecret = config["delegatedUserSecret"];
            //string api = "https://devgcx.sharepoint.com/teams/1000765/_api/site/JoinHubSite('af056a4a-5957-4858-8074-c8fb2e7129fd')";
            //string tenantName = config["tenantName"];
            ////string sharePointUrl = "https://devgcx.sharepoint.com/teams/1000765";

            //ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            //var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
            ////var authManager = new PnP.Framework.AuthenticationManager();
            //var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
            ////var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);


            ////ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            ////var graphClient = new GraphServiceClient(auth);
            ////var site = graphClient.Sites["id"].Request().GetAsync().GetAwaiter().GetResult();

            //try
            //{
            //    HttpClient client = new HttpClient();
            //    client.DefaultRequestHeaders.Add("Authorization", $"Bearer {accessToken}");
            //    var result = client.PostAsync(api, null).Result.EnsureSuccessStatusCode();
            //    log.LogInformation($"result: {result}");
            //}
            //catch (Exception e)
            //{
            //    log.LogError($"Exception getting host: {e.Message}");
            //    if (e.InnerException is not null)
            //        log.LogError($"InnerException: {e.InnerException.Message}");
            //}




            //await HubSiteCSOM.AddToTheHubSiteAsync(log);

            await HubSiteCSOM.SiteToHubAssociation(log);






            return new OkResult();
        }
    }
}

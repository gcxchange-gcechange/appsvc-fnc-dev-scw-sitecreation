using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Online.SharePoint.TenantAdministration;
using Azure.Core;
using Microsoft.Extensions.Configuration;
using PnP.Framework.Diagnostics;
using static appsvc_fnc_dev_scw_sitecreation_dotnet001.Auth;
using Microsoft.Extensions.Logging;
using AngleSharp.Common;
using Microsoft.Graph;

using System.Net.Http;
using System.Net;
using PnP.Framework.Http;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class HubSiteCSOM
    {

        public static async Task JoinTheHubSiteAsync(Microsoft.Extensions.Logging.ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string tenantName = config["tenantName"];
            string sharePointUrl = "https://devgcx.sharepoint.com/teams/1000765";

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
            var authManager = new PnP.Framework.AuthenticationManager();
            var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
            var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

            try {
                var tenant = new Tenant(ctx);
                tenant.ConnectSiteToHubSite("https://devgcx.sharepoint.com/teams/1000765", "https://devgcx.sharepoint.com/sites/communities");
                ctx.ExecuteQuery();
            }
            catch (Exception e)
            {
                log.LogError($"Exception getting host: {e.Message}");
                if (e.InnerException is not null)
                    log.LogError($"InnerException: {e.InnerException.Message}");
            }

            //using (ClientContext context = new ClientContext(tenantSiteUrl))
            //{
            //    SecureString securePassword = new SecureString();
            //    foreach (char c in password.ToCharArray())
            //    {
            //        securePassword.AppendChar(c);
            //    }

            //    context.AuthenticationMode = ClientAuthenticationMode.Default;
            //    context.Credentials = new SharePointOnlineCredentials(userName, securePassword);

            //    var tenant = new Tenant(context);

            //    string hubSiteUrl = "https://tenant.sharepoint.com/sites/HubSiteCollection";
            //    string associateSiteUrl = "https://tenant.sharepoint.com/sites/AssociateSiteCollection";

            //    tenant.ConnectSiteToHubSite(associateSiteUrl, hubSiteUrl);

            //    context.ExecuteQuery();
            //}









        }

        public static async Task AddToTheHubSiteAsync(Microsoft.Extensions.Logging.ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];

            try
            {
                ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
                var graphClient = new GraphServiceClient(auth);
                HttpRequestMessage myrequest = new HttpRequestMessage();
                myrequest.Method = HttpMethod.Post;
                myrequest.RequestUri = new Uri("https://devgcx.sharepoint.com/teams/1000786/_api/site/JoinHubSite('af056a4a-5957-4858-8074-c8fb2e7129fd')");
                var result = await graphClient.HttpProvider.SendAsync(myrequest);
                log.LogInformation($"result = {result}");
            }
            catch (Exception e)
            {
                log.LogError($"Exception: {e.Message}");
                if (e.InnerException is not null)
                    log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                log.LogError($"Source: {e.Source}");
            }
        }



        public static async Task AddToTheHubSiteAsyncTake2(Microsoft.Extensions.Logging.ILogger log)
        {


            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string tenantName = config["tenantName"];
            string sharePointUrl = "https://devgcx.sharepoint.com/teams/1000765";

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
            var authManager = new PnP.Framework.AuthenticationManager();
            var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
            var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);


            HttpWebRequest endpointRequest = (HttpWebRequest)HttpWebRequest.Create("https://devgcx.sharepoint.com/teams/1000786/_api/site/JoinHubSite('af056a4a-5957-4858-8074-c8fb2e7129fd')");
            endpointRequest.Method = "POST";
            endpointRequest.Accept = "application/json;odata=verbose";
            endpointRequest.Headers.Add("Authorization", "Bearer " + accessToken);
            HttpWebResponse endpointResponse = (HttpWebResponse)endpointRequest.GetResponse();

            



        }


        public static async Task SiteToHubAssociation(Microsoft.Extensions.Logging.ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string tenantName = config["tenantName"];
            string sharePointUrl = "https://devgcx.sharepoint.com/teams/1000756";

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
            var authManager = new PnP.Framework.AuthenticationManager();
            var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
            var siteCtx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);


            Guid hubsite = new Guid("af056a4a-5957-4858-8074-c8fb2e7129fd");

            log.LogDebug("Site {siteurl} will be associated with hub {hubsiteID}", siteCtx.Url, hubsite);
            var pnpclient = PnPHttpClient.Instance.GetHttpClient(siteCtx);
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, $"{siteCtx.Url}/_api/site/JoinHubSite('{hubsite.ToString("D")}')")
            {
                Content = null
            };
            request.Headers.Add("accept", "application/json;odata.metadata=none");
            request.Headers.Add("odata-version", "4.0");
            await PnPHttpClient.AuthenticateRequestAsync(request, siteCtx).ConfigureAwait(false);
            HttpResponseMessage response = await pnpclient.SendAsync(request, new System.Threading.CancellationToken());
            if (!response.IsSuccessStatusCode)
                throw new Exception($"Site to hub association failed: {response.StatusCode}");
            log.LogDebug("Site {siteurl} was successfully associated with hub {hubsiteID}", siteCtx.Url, hubsite);
        }





    }
}

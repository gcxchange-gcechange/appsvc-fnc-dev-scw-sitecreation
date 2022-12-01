using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Azure;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Core.Auth;
using PnP.Core.Services;
using PnP.Framework;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using Portable.Xaml.Markup;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;
using ListItem = Microsoft.Graph.ListItem;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    public class CreateSite
    {
        [FunctionName("CreateSite")]
        public void Run([QueueTrigger("sitecreation", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation("CreateSite trigger function processed a request.");

            // todo - get site request data from the queue and replace hard-coded values
            //      - remove uneccessary API permissions
            //      - get userId value
            //      - prevent duplicate calls to create when there is an issue with the queue item (poison)


            dynamic data = JsonConvert.DeserializeObject(myQueueItem);
            log.LogInformation($"myQueueItem: {myQueueItem}");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string aadApplicationId = config["clientId"];
            string description = data?.SpaceDescription;
            string DisplayName = $"{data?.SpaceName} | {data?.SpaceNameFR}";
            string certificateName = config["certificateName"];
            string connectionString = config["AzureWebJobsStorage"];
            string keyVaultUrl = config["keyVaultUrl"];
            string requestId = data?.Id;
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string tenantId = config["tenantId"];
            string userId = config["userId"];

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var teamId = CreateTeam(graphAPIAuth, sharePointUrl, requestId, description, userId, log).GetAwaiter().GetResult();

            if (teamId != string.Empty)
            {
                UpdateName(graphAPIAuth, keyVaultUrl, certificateName, aadApplicationId, tenantId, sharePointUrl, teamId, requestId, DisplayName, log).GetAwaiter().GetResult();
                ApplyTemplate(keyVaultUrl, certificateName, aadApplicationId, sharePointUrl, functionContext, log).GetAwaiter().GetResult();
                _ = AddOwner(graphAPIAuth, teamId, "e4b36075-bb6a-4acf-badb-076b0c3d8d90", log);
                _ = AddToQueue(connectionString, requestId, DisplayName, (string)data?.RequesterName, (string)data?.RequesterEmail, log);

                // need to get trigger Unclassified via queue - LIST
                // remove user RemoveOwner - call this only if label has been successfully applied
                // also do research about await






            }
            else
            {
                log.LogInformation("Site already exists");
            }
        }

        public static async Task<string> CreateTeam(GraphServiceClient graphClient, string sharePointUrl, string requestId, string description, string userId, ILogger log)
        {
            log.LogInformation("CreateTeam received a request.");

            var teamId = string.Empty;

            // make sure team site does not already exist
            System.Net.Http.HttpClient client = new System.Net.Http.HttpClient();
            var response = await client.GetAsync(sharePointUrl);
            if (response.StatusCode != System.Net.HttpStatusCode.NotFound)
            {
                return string.Empty;
            }

            try
            {
                var team = new Team
                {
                    Description = description,
                    DisplayName = requestId,
                    Members = new TeamMembersCollectionPage()
                {
                    new AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {
                            "owner"
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userId}')"}
                        }
                    }
                },
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
                    }
                };

                var teamResponse = await graphClient.Teams.Request().AddResponseAsync(team);
                if (teamResponse.HttpHeaders.TryGetValues("Location", out var headerValues))
                {
                    teamId = headerValues?.First().Split('\'', StringSplitOptions.RemoveEmptyEntries)[1];
                }
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                return string.Empty;
            }

            log.LogInformation($"CreateTeam received a request. teamId: {teamId}");
            return teamId;
        }

        public static async Task<IActionResult> AddToQueue(string connectionString, string requestId, string displayName, string requesterName, string requesterEmail, ILogger log)
        {
            log.LogInformation("AddToQueue received a request.");

            try
            {
                // send item to email queue to trigger email to user
                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                        {"Title", displayName},
                        {"RequesterName", requesterName},
                        {"RequesterEmail", requesterEmail},
                        {"Status", "Team Created"},
                        {"Comment", ""}
                        }
                    }
                };

                listItem.Id = requestId;
                Common.InsertMessageAsync(connectionString, "email", listItem, log).GetAwaiter().GetResult();
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                return new BadRequestResult();
            }

            log.LogInformation("AddToQueue processed a request.");
            return new OkResult();
        }

        public static async Task<IActionResult> AddOwner(GraphServiceClient graphClient, string teamId, string userId, ILogger log)
        {
            log.LogInformation("AddOwner received a request.");

            try {
                var values = new List<ConversationMember>()
                {
                    new AadUserConversationMember
                    {
                        Roles = new List<String>() { "owner" },
                        AdditionalData = new Dictionary<string, object>() { {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userId}')"} }
                    }
                };

                await graphClient.Teams[teamId].Members.Add(values).Request().PostAsync();
            }
            catch (Exception e) 
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                //return new BadRequestResult();
            }

            log.LogInformation("AddOwner processed a request.");
            return new OkResult();
        }

        public static async Task<IActionResult> RemoveOwner(GraphServiceClient graphClient, string teamId, string userId, ILogger log)
        {
            log.LogInformation("RemoveOwner received a request.");

            try
            {
                await graphClient.Teams[teamId].Members[userId].Request().DeleteAsync();
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                //return new BadRequestResult();
            }

            log.LogInformation("RemoveOwner processed a request.");
            return new OkResult();
        }


        public static async Task<IActionResult> UpdateName(GraphServiceClient graphClient, string keyVaultUrl, string certificateName, string aadApplicationId, string tenantId, string sharePointUrl, string teamId, string requestId, string displayName, ILogger log)
        {
            log.LogInformation("UpdateName received a request.");

            log.LogInformation($"Update Team where teamId is {teamId}");
            try
            {
                var team = new Team { DisplayName = displayName };
                await graphClient.Teams[teamId].Request().UpdateAsync(team);
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
            }

            log.LogInformation($"Update Share Point where sharePointUrl is {sharePointUrl}");
            try
            {
                // Get certificate from the key vault
                X509Certificate2 mycert = await Auth.GetKeyVaultCertificateAsync(keyVaultUrl, certificateName, log);

                // Creates and configures the host
                var host = Host.CreateDefaultBuilder()
                    .ConfigureServices((context, services) =>
                    {
                        // Add PnP Core SDK
                        services.AddPnPCore(options =>
                        {
                            // Configure the interactive authentication provider as default
                            options.DefaultAuthenticationProvider = new X509CertificateAuthenticationProvider(aadApplicationId, tenantId, mycert);
                        });
                    })
                    .UseConsoleLifetime()
                    .Build();

                // Start the host
                await host.StartAsync();

                using (var scope = host.Services.CreateScope())
                {
                    // Ask an IPnPContextFactory from the host
                    var pnpContextFactory = scope.ServiceProvider.GetRequiredService<IPnPContextFactory>();

                    // Create a PnPContext
                    using (var context = await pnpContextFactory.CreateAsync(new Uri(sharePointUrl)))
                    {
                        context.GraphFirst = false; // reference: https://pnp.github.io/pnpcore/using-the-sdk/basics-apis.html
                        var web = await context.Web.GetAsync(p => p.Title);
                        web.Title = displayName;
                        await web.UpdateAsync();
                    }
                }
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
            }

            //log.LogInformation("Update Group");
            //try
            //{
            //    var group = new Group { DisplayName = displayName };
            //    await graphClient.Groups[teamId].Request().UpdateAsync(group);
            //}
            //catch (Exception e)
            //{
            //    log.LogInformation($"Message: {e.Message}");
            //    if (e.InnerException is not null)
            //        log.LogInformation($"InnerException: {e.InnerException.Message}");
            //}

            log.LogInformation("UpdateName processed a request.");

            return new OkResult();
        }

        public static async Task<IActionResult> ApplyTemplate(string keyVaultUrl, string certificateName, string aadApplicationId, string sharePointUrl, ExecutionContext functionContext, ILogger log)
        {
            log.LogInformation("ApplyTemplate received a request.");

            X509Certificate2 mycert = await Auth.GetKeyVaultCertificateAsync(keyVaultUrl, certificateName, log);

            string tenantName = "devgcx";

            AuthenticationManager auth = new AuthenticationManager(aadApplicationId, mycert, $"{tenantName}.onmicrosoft.com");
            ClientContext ctx = await auth.GetContextAsync(sharePointUrl);

            try
            {

                ctx.RequestTimeout = Timeout.Infinite;
                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();

                log.LogInformation($"Successfully connected to site: {web.Title}");

                DirectoryInfo dInfo;
                var schemaDir = "";

                string currentDirectory = functionContext.FunctionDirectory;

                if (currentDirectory == null)
                {
                    string workingDirectory = Environment.CurrentDirectory;
                    currentDirectory = System.IO.Directory.GetParent(workingDirectory).Parent.Parent.FullName;
                    dInfo = new DirectoryInfo(currentDirectory);
                    schemaDir = dInfo + "\\GxDcCPS-SitesCreations-fnc\\bin\\Debug\\net461\\Templates\\GenericTemplate";
                }
                else
                {
                    dInfo = new DirectoryInfo(currentDirectory);
                    schemaDir = dInfo.Parent.FullName + "\\Templates\\GenericTemplate";
                }

               DirectoryInfo dInfo2 = new DirectoryInfo(schemaDir);

               XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");

                string PNP_TEMPLATE_FILE = "template-name.xml";
                ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);
                log.LogInformation($"Successfully found template with ID '{template.Id}'");

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
                    }
                };
                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");

                template.Connector = connector;

                log.LogInformation("ApplyProvisioningTemplate...");
                web.ApplyProvisioningTemplate(template, ptai);
                log.LogInformation("...worked!");

                log.LogInformation($"Site {web.Title} apply template successfully.");
            }
            catch (ReflectionTypeLoadException ex)
            {
                foreach (var item in ex.LoaderExceptions)
                {
                    log.LogInformation(item.Message);
                }
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");

            }

            log.LogInformation("ApplyTemplate processed a request.");
            return new OkResult();
        }

        // log.LogInformation("UpdateName processed a request.");
        // var group = new Group { DisplayName = displayName };
        // var team = new Team { DisplayName = displayName };

        // // recollection before testing - OP, 2022.11.03
        // // these two methods appear to update AAD at the same time
        // // but not SharePoint

        // //await graphClient.Groups[teamId].Request().UpdateAsync(group); // updated aad immediately, still not on SharePoint after an hour
        // await graphClient.Teams[teamId].Request().UpdateAsync(team); // updated aad (almost) immediately, updated SharePoint in minutes -- weird, I thought it took longer last time!!
        //                                                              // ok ran again and did not update SP after about an hour
        // // method 3: updates SharePoint URL almost immediately; aad not so much... ; actually SP Admin got the update before aad, still waiting... 26 minutes later
        //// SharePoint Admin Centre has it's own lag issues
    }
}
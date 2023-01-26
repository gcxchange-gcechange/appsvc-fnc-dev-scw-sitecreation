using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using AuthenticationManager = PnP.Framework.AuthenticationManager;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;
using ListItem = Microsoft.Graph.ListItem;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    public class CreateSite
    {
        [FunctionName("CreateSite")]
        public static async Task RunAsync([QueueTrigger("sitecreation", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation("CreateSite trigger function received a request.");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string aadApplicationId = config["clientId"];
            string certificateName = config["certificateName"];
            string connectionString = config["AzureWebJobsStorage"];
            string description = data?.SpaceDescription;
            string DisplayName = $"{data?.SpaceName} | {data?.SpaceNameFR}";
            string keyVaultUrl = config["keyVaultUrl"];
            string ownerId = config["ownerId"];
            string queueName = data?.SecurityCategory;
            string requestId = data?.Id;

            int newRequestId = Int32.Parse(requestId) + 500;    // offset to ensure unique Id
            requestId = newRequestId.ToString();

            string RequesterEmail = data?.RequesterEmail;
            string RequesterName = data?.RequesterName;
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string userId = config["userId"];

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            var groupId = await CreateGroup(graphClient, sharePointUrl, requestId, DisplayName, description, userId, log);

            if (groupId != string.Empty)
            {
                await AddLicensedUserToGroup(graphClient, log, groupId, userId, ownerId);

                // wait 3 minutes to allow for provisioning
                Thread.Sleep(3 * 60 * 1000);

                await AddTeam(graphClient, log, groupId);
                await ApplyTemplate(keyVaultUrl, certificateName, aadApplicationId, sharePointUrl, DisplayName, functionContext, log);
                await AddToSensitivityQueue(connectionString, queueName, requestId, groupId, DisplayName, RequesterName, RequesterEmail, log);
            }
            else
            {
                log.LogInformation("Site already exists");
            }

            log.LogInformation("CreateSite trigger function processed a request.");
        }

        public static async Task<bool> AddToSensitivityQueue(string connectionString, string queueName, string requestId, string groupId, string DisplayName, string RequesterName, string RequesterEmail, ILogger log)
        {
            log.LogInformation("AddToSensitivityQueue received a request.");

            ListItem listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"Id", requestId},
                        {"groupId", groupId},
                        {"DisplayName", DisplayName},
                        {"RequesterName", RequesterName},
                        {"RequesterEmail", RequesterEmail}
                    }
                }
            };

            await Common.InsertMessageAsync(connectionString, queueName, listItem, log);

            log.LogInformation("AddToSensitivityQueue processed a request.");

            return true;
        }

        public static async Task<string> CreateGroup(GraphServiceClient graphClient, string sharePointUrl, string requestId, string displayName, string description, string userId, ILogger log)
        {
            log.LogInformation($"CreateTeam received a request. requestId: {requestId}");
            log.LogInformation($"sharePointUrl: {sharePointUrl}");

            // make sure team site does not already exist
            HttpClient client = new HttpClient();
            var response = await client.GetAsync(sharePointUrl);
            if (response.StatusCode != HttpStatusCode.NotFound)
                return string.Empty;

            string groupId;

            try
            {
                log.LogInformation($"create group obj");
                var o365Group = new Microsoft.Graph.Group
               {
                   Description = description,
                   DisplayName = $@"{displayName}",
                   GroupTypes = new List<String>() { "Unified" },
                   MailEnabled = true,
                   MailNickname = requestId,
                   SecurityEnabled = false,
                   Visibility = "Private"
               };

                log.LogInformation($"pre-AddASync");
                var result = await graphClient.Groups.Request().AddAsync(o365Group);
                log.LogInformation($"post-AddASync");
                groupId = result.Id;
                log.LogInformation($"Site and Office 365 {displayName} created successfully. And groupId: {groupId}");
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");

                groupId = string.Empty;
            }

            log.LogInformation($"CreateTeam processed a request. groupId: {groupId}");

            return groupId;
        }

        public static async Task<bool> AddLicensedUserToGroup(GraphServiceClient graphClient, ILogger log, string groupId, string TEAMS_INIT_USERID, string ownerId)
        {
            try {
                var directoryObject = new DirectoryObject { Id = TEAMS_INIT_USERID }; //teamcreator
                await graphClient.Groups[groupId].Owners.References.Request().AddAsync(directoryObject);

                directoryObject = new DirectoryObject { Id = ownerId };
                await graphClient.Groups[groupId].Owners.References.Request().AddAsync(directoryObject);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation($"Licensed add to owner of {groupId} successfully.");

            return true;
        }

        public static async Task<bool> AddTeam(GraphServiceClient graphClient, ILogger log, string groupId)
        {
            log.LogInformation($"---");
            log.LogInformation($"Add Team to {groupId}.");
            log.LogInformation($"---");

            try {
                var team = new Team
                {
                    MemberSettings = new TeamMemberSettings
                    {
                        AllowCreateUpdateChannels = true
                    },
                    MessagingSettings = new TeamMessagingSettings
                    {
                        AllowUserEditMessages = true,
                        AllowUserDeleteMessages = true
                    },
                    FunSettings = new TeamFunSettings
                    {
                        AllowGiphy = true,
                        GiphyContentRating = GiphyRatingType.Strict
                    }
                };

               await graphClient.Groups[groupId].Team.Request().PutAsync(team);
            }
            catch (Exception e)
            {
                log.LogInformation($"---");
                log.LogInformation($"Team creation failed!!");
                log.LogInformation($"---");

                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation($"---");
            log.LogInformation($"Team created successfully.");
            log.LogInformation($"---");

            return true;
        }

        public static async Task<bool> ApplyTemplate(string keyVaultUrl, string certificateName, string aadApplicationId, string sharePointUrl, string DisplayName, ExecutionContext functionContext, ILogger log)
        {
            log.LogInformation($"---");
            log.LogInformation("ApplyTemplate received a request.");
            log.LogInformation($"---");

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
                log.LogInformation($"---");

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
                log.LogInformation($"---");

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
                    }
                };

                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");
                template.Connector = connector;

                //template.Parameters.Add("DisplayName", "Dispay Name Test");
                //template.Parameters.Add("Title", "Title Test");

                log.LogInformation("ApplyProvisioningTemplate...");
                try
                {
                    web.ApplyProvisioningTemplate(template, ptai);
                }
                catch (Exception e)
                {
                    log.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                    log.LogError($"StackTrace: {e.StackTrace}");
                }
                log.LogInformation("...worked!");

                log.LogInformation($"Site {web.Title} apply template successfully.");
                log.LogInformation($"---");
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
            log.LogInformation($"---");

            return true;
        }
    }
}
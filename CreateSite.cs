using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.News.DataModel;
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
            //string ownerId = config["ownerId"];
            string owners = data?.Owner1;
            string queueName = data?.SecurityCategory;
            string requestId = data?.Id;

            int newRequestId = Int32.Parse(requestId) + 500;    // offset to ensure unique Id
            requestId = newRequestId.ToString();

            string RequesterEmail = data?.RequesterEmail;
            string RequesterName = data?.RequesterName;
            string sharePointUrl = config["sharePointUrl"] + requestId;
            string userId = config["userId"];

            string members = data?.Members;

            string siteId = config["siteId"];
            string listId = config["listId"];
            string itemId = data?.Id;

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            var groupId = await CreateGroup(graphClient, sharePointUrl, requestId, DisplayName, description, userId, log);
            log.LogInformation($"teamId: {groupId}");

            if (groupId != string.Empty)
            {
                await UpdateSiteUrl(graphClient, sharePointUrl, siteId, listId, itemId, log);

                await AddOwnersToGroup(graphClient, log, groupId, userId, owners);

                // wait 3 minutes to allow for provisioning
                Thread.Sleep(3 * 60 * 1000);

                var teamId = await AddTeam(graphClient, log, groupId);
                log.LogInformation($"teamId: {teamId}");

                await ApplyTemplate(keyVaultUrl, certificateName, aadApplicationId, sharePointUrl, DisplayName, functionContext, log);

                await AddMembersToTeam(graphClient, log, groupId, teamId, members);

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


        public static async Task<string> UpdateSiteUrl(GraphServiceClient graphClient, string sharePointUrl, string siteId, string listId, string itemId, ILogger log)
        {
            log.LogInformation("UpdateSiteUrl received a request.");

            log.LogInformation($"sharePointUrl: {sharePointUrl}");
            log.LogInformation($"siteId: {siteId}");
            log.LogInformation($"listId: {listId}");
            log.LogInformation($"itemId: {itemId}");

            try
            {
                var fieldValueSet = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"SiteUrl", sharePointUrl}
                    }
                };

                await graphClient.Sites[siteId].Lists[listId].Items[itemId].Fields.Request().UpdateAsync(fieldValueSet);
            }

            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("UpdateSiteUrl processed a request.");

            return string.Empty;
        }

        public static async Task<string> CreateGroup(GraphServiceClient graphClient, string sharePointUrl, string requestId, string displayName, string description, string userId, ILogger log)
        {
            log.LogInformation($"CreateGroup received a request. requestId: {requestId}");
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

                var result = await graphClient.Groups.Request().AddAsync(o365Group);
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

            log.LogInformation($"CreateGroup processed a request. groupId: {groupId}");

            return groupId;
        }

        public static async Task<bool> AddOwnersToGroup(GraphServiceClient graphClient, ILogger log, string groupId, string TEAMS_INIT_USERID, string owners)
        {
            try {
                var directoryObject = new DirectoryObject { Id = TEAMS_INIT_USERID }; //teamcreator
                await graphClient.Groups[groupId].Owners.References.Request().AddAsync(directoryObject);

                foreach (string id in owners.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    directoryObject = new DirectoryObject { Id = id };
                    await graphClient.Groups[groupId].Owners.References.Request().AddAsync(directoryObject);
                }
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

        public static async Task<bool> AddMembersToTeam(GraphServiceClient graphClient, ILogger log, string groupId, string teamId, string Members)
        {
            log.LogInformation("AddMembersToTeam received a request.");

            try
            {
                foreach (string memberId in Members.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    log.LogInformation($"memberId: {memberId}");

                    var directoryObject = new DirectoryObject
                    {
                        Id = memberId
                    };
                    await graphClient.Groups[groupId].Members.References.Request().AddAsync(directoryObject);

                    AadUserConversationMember mem = new AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {
                            "member"
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{memberId}')"}
                        }
                    };
                    await graphClient.Teams[teamId].Members.Request().AddAsync(mem);
                }










            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

           log.LogInformation("AddMembersToTeam processed a request.");

            return true;
        }

        public static async Task<string> AddTeam(GraphServiceClient graphClient, ILogger log, string groupId)
        {
            log.LogInformation("AddTeam received a request.");

            string teamId = string.Empty;

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

                var t = await graphClient.Groups[groupId].Team.Request().PutAsync(team);
                teamId = t.Id;
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

            log.LogInformation("AddTeam processed a request.");

            return teamId;
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

                string PNP_TEMPLATE_FILE = "template-new.xml";

                ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);
                log.LogInformation($"Successfully found template with ID '{template.Id}'");


                //ContentTypeBinding bindingToLookFor = new ContentTypeBinding();
                //bindingToLookFor.ContentTypeId = "0x01002CF74A4DAE39480396EEA7A4BA2BE5FB";

                ////var offendingLists = template.Lists.Where(l => l.ContentTypeBindings.Contains(bindingToLookFor));
                //var offendingLists = template.Lists;

                //log.LogInformation("Look for content binding...");
                //foreach (var l in offendingLists)
                //{
                //    //l.ContentTypeBindings.Remove(bindingToLookFor);
                //    log.LogInformation($"l.Title: {l.Title} - l.Description: {l.Description}");

                //    log.LogInformation("Bindings:");
                //    foreach (var b in l.ContentTypeBindings)
                //    {
                //        log.LogInformation($"b.ContentTypeId: {b.ContentTypeId}");

                //        if (b.ContentTypeId == "0x01002CF74A4DAE39480396EEA7A4BA2BE5FB")
                //        {
                //            log.LogInformation("Found it!");
                //            l.ContentTypeBindings.Remove(b);
                //            // Message: Collection was modified; enumeration operation may not execute.
                //        }


                //    }
                //    log.LogInformation("----------------------------------------------------");
                //}

                //log.LogInformation("...done?");


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
                //
                //log.LogInformation("SiteGroups:");
                //foreach (var g in template.Security.SiteGroups)
                //{
                //    log.LogInformation($"g.Description: {g.Description}");
                //}
                //template.Security.AdditionalAdministrators.Add()
                // template.Security.SiteGroups[].Description

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
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using PnP.Framework.Http;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using static appsvc_fnc_dev_scw_sitecreation_dotnet001.Auth;
using ExecutionContext = Microsoft.Azure.WebJobs.ExecutionContext;
using ListItem = Microsoft.Graph.ListItem;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    public class CreateSite
    {
        private static string teamsUrl = string.Empty;

        [FunctionName("CreateSite")]
        public static async Task RunAsync([QueueTrigger("sitecreation", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation("CreateSite trigger function received a request.");

            // assign variables from config
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string connectionString = config["AzureWebJobsStorage"];
            string creatorId = config["ownerId"];
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string followingContentFeatureId = config["followingContentFeatureId"];
            string listId = config["listId"];
            string siteId = config["siteId"];
            //string teamsChannelId = config["teamsChannelId"];
            string tenantId = config["tenantId"];
            string tenantName = config["tenantName"];

            // assign variables from queue
            dynamic data = JsonConvert.DeserializeObject(myQueueItem);
            string descriptionEn = data?.SpaceDescription;
            string descriptionFr = data?.SpaceDescriptionFR;
            string displayName = $"{data?.SpaceName} - {data?.SpaceNameFR}";
            string itemId = data?.Id;
            string owners = data?.Owner1;
            string queueName = data?.SecurityCategory;
            string requesterEmail = data?.RequesterEmail;
            string requesterName = data?.RequesterName;

            // manipulated values
            // - take id from SharePoint list and append prefix to use as part of url
            string sitePath = string.Concat("1000", itemId);
            string sharePointUrl = string.Concat(config["sharePointUrl"], sitePath);

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            var groupId = await CheckAndCreateGroup(graphClient, sharePointUrl, sitePath, displayName, descriptionEn, creatorId, owners, log);

            if (groupId != string.Empty)
            {
                await UpdateSiteUrl(delegatedUserName, delegatedUserSecret, sharePointUrl, siteId, listId, itemId, log);

                //await AddOwnersToGroup(graphClient, log, groupId, creatorId, owners);

                // wait 3 minutes to allow for provisioning
                Thread.Sleep(3 * 60 * 1000);

                var teamId = await AddTeam(groupId, tenantId, delegatedUserName, delegatedUserSecret, log);

                await ApplyTemplate(sharePointUrl, tenantName, descriptionEn, descriptionFr, followingContentFeatureId, teamsUrl, delegatedUserName, delegatedUserSecret, functionContext, log);

                // deferred functionality
                //await AddMembersToTeam(graphClient, log, groupId, teamId, members);

                await AddToSensitivityQueue(connectionString, queueName, itemId, sitePath, groupId, displayName, requesterName, requesterEmail, log);
            }
            else
            {
                log.LogInformation("Site already exists");
            }

            log.LogInformation("CreateSite trigger function processed a request.");
        }

        public static async Task<bool> AddToSensitivityQueue(string connectionString, string queueName, string itemId, string sitePath, string groupId, string DisplayName, string RequesterName, string RequesterEmail, ILogger log)
        {
            log.LogInformation("AddToSensitivityQueue received a request.");

            ListItem listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"Id", sitePath},
                        {"itemId", itemId},
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

        public static async Task<string> UpdateSiteUrl(string userName, string userSecret, string sharePointUrl, string siteId, string listId, string itemId, ILogger log)
        {
            log.LogInformation("UpdateSiteUrl received a request.");

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(userName, userSecret, log);
            var graphClient = new GraphServiceClient(auth);

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

        public static async Task<string> CheckAndCreateGroup(GraphServiceClient graphClient, string sharePointUrl, string sitePath, string displayName, string description, string creatorId, string owners, ILogger log)
        {
            log.LogInformation($"CreateGroup received a request.");
            log.LogInformation($"sharePointUrl: {sharePointUrl}");

            // make sure team site does not already exist
            HttpClient client = new HttpClient();
            var response = await client.GetAsync(sharePointUrl);
            //either option not making a difference: HttpCompletionOption.ResponseHeadersRead
            log.LogInformation($"response.StatusCode: {response.StatusCode}");
            if (response.StatusCode != HttpStatusCode.NotFound && response.StatusCode != HttpStatusCode.Forbidden)
                return string.Empty;

            string groupId;

            try
            {
                List<string> ownerList = new List<string>
                {
                    $"https://graph.microsoft.com/v1.0/users/{creatorId}"
                };

                foreach (string email in owners.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var user = await graphClient.Users[email].Request().GetAsync();
                    ownerList.Add($"https://graph.microsoft.com/v1.0/users/{user.Id}");
                }

                var o365Group = new Microsoft.Graph.Group
                {
                    Description = description,
                    DisplayName = $@"{displayName}",
                    GroupTypes = new List<String>() { "Unified" },
                    MailEnabled = true,
                    MailNickname = sitePath,
                    SecurityEnabled = false,
                    Visibility = "Private",
                    AdditionalData = new Dictionary<string, object>
                    {
                        {"owners@odata.bind" , ownerList}
                    }               
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

        //public static async Task<bool> AddOwnersToGroup(GraphServiceClient graphClient, ILogger log, string groupId, string creatorId, string owners)
        //{
        //    log.LogInformation("AddOwnersToGroup received a request.");

        //    try
        //    {
        //        await graphClient.Groups[groupId].Owners.References.Request().AddAsync(new DirectoryObject { Id = creatorId });

        //        foreach (string email in owners.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
        //        {
        //            var user = await graphClient.Users[email].Request().GetAsync();
        //            var id = user.Id;
        //            await graphClient.Groups[groupId].Owners.References.Request().AddAsync(new DirectoryObject { Id = id });
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        log.LogError($"Message: {e.Message}");
        //        if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
        //        log.LogError($"StackTrace: {e.StackTrace}");
        //    }

        //    log.LogInformation("AddOwnersToGroup processed a request.");

        //    return true;
        //}

        public static async Task<bool> AddMembersToTeam(GraphServiceClient graphClient, ILogger log, string groupId, string teamId, string Members)
        {
            log.LogInformation("AddMembersToTeam received a request.");

            try
            {
                foreach (string email in Members.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var user  = await graphClient.Users[email].Request().GetAsync();
                    var memberId = user.Id;

                    log.LogInformation($"email: {email}");
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

        public static async Task<string> AddTeam(string groupId, string tenantId, string userName, string userSecret, Microsoft.Extensions.Logging.ILogger log)
        {
            log.LogInformation("AddTeam received a request.");

            string teamId = string.Empty;

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(userName, userSecret, log);
            var graphClient = new GraphServiceClient(auth);

            try
            {
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

                var channels = await graphClient.Teams[teamId].Channels.Request().GetAsync();
                var channelId = "";
                foreach (var channel in channels)
                {
                    channelId = channel.Id;
                }
                teamsUrl = $@"https://teams.microsoft.com/l/team/{channelId}/conversations?groupId={teamId}&tenantId={tenantId}";
                log.LogInformation($"teamsUrl = {teamsUrl}");
            }
            catch (Exception e)
            {
                log.LogInformation("Team creation failed!!");

                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation($"AddTeam processed a request. teamId: {teamId}");

            return teamId;
        }

        public static async Task<bool> ApplyTemplate(string sharePointUrl, string tenantName, string descriptionEn, string descriptionFr, string followingContentFeatureId, string teamsUrl, string userName, string userSecret, ExecutionContext functionContext, ILogger log)
        {
            log.LogInformation("ApplyTemplate received a request.");

            try
            {
                ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(userName, userSecret, log);
                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                //IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
                //string aadApplicationId = config["clientId"];
                //string certificateName = config["certificateName"];
                //string keyVaultUrl = config["keyVaultUrl"];

                //X509Certificate2 mycert = await Auth.GetKeyVaultCertificateAsync(keyVaultUrl, certificateName, log);
                //AuthenticationManager auth = new AuthenticationManager(aadApplicationId, mycert, $"{tenantName}.onmicrosoft.com");
                //ClientContext ctx = await auth.GetContextAsync(sharePointUrl);

                Web web = ctx.Web;
                ctx.Load(web, w => w.Title);
                ctx.ExecuteQuery();

                log.LogInformation($"Successfully connected to site: {web.Title}");

                //deactivate the following content feature
                web.DeactivateFeature(Guid.Parse(followingContentFeatureId));

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

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
                    }//,
                    //PersistTemplateInfo = true
                    //IgnoreDuplicateDataRowErrors
                    //ClearNavigation
                };

                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");
                template.Connector = connector;

                template.Parameters.Add("DescriptionEn", descriptionEn);
                template.Parameters.Add("DescriptionFr", descriptionFr);
                template.Parameters.Add("HubSiteUrl", "https://devgcx.sharepoint.com/sites/communities");
                template.Parameters.Add("MSTeamsUrl", teamsUrl);

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
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogError($"InnerException: {e.InnerException.Message}");

            }

            log.LogInformation("ApplyTemplate processed a request.");

            return true;
        }


        public static async Task SiteToHubAssociation(string sharePointUrl, string hubSiteId, ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string tenantName = config["tenantName"];
            //string sharePointUrl = "https://devgcx.sharepoint.com/teams/1000756";

            ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
            var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
            var authManager = new PnP.Framework.AuthenticationManager();
            var accessToken = await auth.GetTokenAsync(new TokenRequestContext(scopes), new System.Threading.CancellationToken());
            var siteCtx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);


            //Guid hubsite = new Guid("af056a4a-5957-4858-8074-c8fb2e7129fd");

            log.LogDebug("Site {siteurl} will be associated with hub {hubsiteID}", siteCtx.Url, hubSiteId);
            var pnpclient = PnPHttpClient.Instance.GetHttpClient(siteCtx);
            HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, $"{siteCtx.Url}/_api/site/JoinHubSite('{hubSiteId}')")
            {
                Content = null
            };
            request.Headers.Add("accept", "application/json;odata.metadata=none");
            request.Headers.Add("odata-version", "4.0");
            await PnPHttpClient.AuthenticateRequestAsync(request, siteCtx).ConfigureAwait(false);
            HttpResponseMessage response = await pnpclient.SendAsync(request, new System.Threading.CancellationToken());
            if (!response.IsSuccessStatusCode)
                throw new Exception($"Site to hub association failed: {response.StatusCode}");
            log.LogDebug("Site {siteurl} was successfully associated with hub {hubsiteID}", siteCtx.Url, hubSiteId);
        }



    }
}
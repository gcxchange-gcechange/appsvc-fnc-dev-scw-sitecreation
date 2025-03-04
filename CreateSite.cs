using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Microsoft.AspNetCore.Mvc;
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
using ILogger = Microsoft.Extensions.Logging.ILogger;
using ListItem = Microsoft.Graph.ListItem;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    public class CreateSite
    {
        private static string teamsUrl = string.Empty;

        public enum GroupCreationStatus : int
        {
            Success = 0,
            SiteExists = 1,
            NoOwner = 2,
            Unknown = 3
        }

        [FunctionName("CreateSite")]
        public static async Task RunAsync([QueueTrigger("sitecreation", Connection = "AzureWebJobsStorage")] string myQueueItem, ILogger log, ExecutionContext functionContext)
        {
            log.LogInformation("CreateSite trigger function received a request.");

            // assign variables from config
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();
            string apprefSiteId = config["apprefSiteId"];
            string connectionString = config["AzureWebJobsStorage"];
            string creatorId = config["ownerId"];
            string delegatedUserName = config["delegatedUserName"];
            string delegatedUserSecret = config["delegatedUserSecret"];
            string followingContentFeatureId = config["followingContentFeatureId"];
            string hubSiteId = config["hubSiteId"];
            string listId = config["listId"];
            string siteId = config["siteId"];
            string teamsLinkListId = config["teamsLinkListId"];
            string tenantId = config["tenantId"];
            string tenantName = config["tenantName"];

            // assign variables from queue
            dynamic data = JsonConvert.DeserializeObject(myQueueItem);
            string descriptionEn = data?.SpaceDescription;
            string descriptionFr = data?.SpaceDescriptionFR;
            string itemId = data?.Id;
            string owners = data?.Owner1;
            string queueName = data?.SecurityCategory;
            string requesterEmail = data?.RequesterEmail;
            string requesterName = data?.RequesterName;
            string SpaceNameEn = data?.SpaceName;
            string SpaceNameFr = data?.SpaceNameFR;

            // manipulated values
            string description = $"{descriptionEn} - {descriptionFr}";
            string displayName = $"{SpaceNameEn} - {SpaceNameFr}";
            string sitePath = string.Concat("1000", itemId);

            if (queueName == "prob")
            {
                // prepend sharePointUrl with "b"
                sitePath = string.Concat("b", sitePath);
            }

            string sharePointUrl = string.Concat(config["sharePointUrl"], sitePath);

            Auth auth = new Auth();
            var graphClient = auth.graphAuth(log);

            var groupId = await CheckAndCreateGroup(graphClient, sharePointUrl, sitePath, displayName, description, creatorId, owners, log);

            Guid guidOutput;
            GroupCreationStatus status;
            bool isValid = Guid.TryParse(groupId, out guidOutput);

            if (isValid)
                status = GroupCreationStatus.Success;
            else
                status = Enum.Parse<GroupCreationStatus>(groupId);

            log.LogInformation($"Group creation status: {status}");

            if (status == GroupCreationStatus.Success)
            {
                ROPCConfidentialTokenCredential tokenCredential = new ROPCConfidentialTokenCredential(delegatedUserName, delegatedUserSecret, log);
                var scopes = new string[] { $"https://{tenantName}.sharepoint.com/.default" };
                var authManager = new PnP.Framework.AuthenticationManager();
                var accessToken = await tokenCredential.GetTokenAsync(new TokenRequestContext(scopes), new CancellationToken());
                var ctx = authManager.GetAccessTokenContext(sharePointUrl, accessToken.Token);

                await UpdateSiteUrl(tokenCredential, sharePointUrl, siteId, listId, itemId, log);

                // wait 3 minutes to allow for provisioning
                Thread.Sleep(3 * 60 * 1000);

                var teamsId = await AddTeam(groupId, tenantId, delegatedUserName, delegatedUserSecret, log);

                await AddToTeamsLinkList(tokenCredential, apprefSiteId, teamsLinkListId, displayName, teamsId, teamsUrl, log);

                await SiteToHubAssociation(ctx, hubSiteId, log);

                await ApplyTemplate(ctx, queueName, descriptionEn, descriptionFr, followingContentFeatureId, teamsUrl, functionContext, log);

                // deferred functionality
                //await AddMembersToTeam(graphClient, log, groupId, teamsId, members);

                await AddToSensitivityQueue(connectionString, queueName, itemId, sitePath, groupId, SpaceNameEn, SpaceNameFr, requesterName, requesterEmail, log);
            }
            else if (status == GroupCreationStatus.SiteExists)
            {
                await AddToStatusQueue(connectionString, itemId, "Site Exists", log);
            }
            else if (status == GroupCreationStatus.NoOwner)
            {
                await AddToStatusQueue(connectionString, itemId, "No Owner", log);
            }

            log.LogInformation("CreateSite trigger function processed a request.");
        }

        public static async Task<bool> AddToSensitivityQueue(string connectionString, string queueName, string itemId, string sitePath, string groupId, string SpaceNameEn, string SpaceNameFr, string RequesterName, string RequesterEmail, ILogger log)
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
                        {"SpaceName", SpaceNameEn},
                        {"SpaceNameFR", SpaceNameFr},
                        {"RequesterName", RequesterName},
                        {"RequesterEmail", RequesterEmail}
                    }
                }
            };

            await Common.InsertMessageAsync(connectionString, queueName, listItem, log);

            log.LogInformation("AddToSensitivityQueue processed a request.");

            return true;
        }

        public static async Task<string> UpdateSiteUrl(ROPCConfidentialTokenCredential tokenCredential, string sharePointUrl, string siteId, string listId, string itemId, ILogger log)
        {
            log.LogInformation("UpdateSiteUrl received a request.");

            var graphClient = new GraphServiceClient(tokenCredential);

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
            log.LogInformation($"response.StatusCode: {response.StatusCode}");

            if (response.StatusCode != HttpStatusCode.NotFound && response.StatusCode != HttpStatusCode.Forbidden)
                return Convert.ToString(GroupCreationStatus.SiteExists);

            string groupId;

            try
            {
                List<string> ownerList = new List<string>
                {
                    $"https://graph.microsoft.com/v1.0/users/{creatorId}"
                };

                List<string> memberList = new List<string>();

                foreach (string email in owners.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
                {
                    log.LogInformation($"email = {email}");

                    try
                    {
                        var user = await graphClient.Users.Request().Filter(Uri.EscapeDataString($"mail eq '{email.Trim().Replace("'", "''")}'")).GetAsync();

                        if (user != null)
                        {
                            string Id = user[0].Id;
                            log.LogInformation($"Id = {Id}");
                            ownerList.Add($"https://graph.microsoft.com/v1.0/users/{Id}");
                            memberList.Add($"https://graph.microsoft.com/v1.0/users/{Id}");
                        }
                        else
                        {
                            log.LogInformation($"Id not found for user {email}");
                        }
                    }
                    catch (Exception e)
                    {
                        log.LogInformation($"Id not found for user {email}");
                        log.LogError($"Message: {e.Message}");
                        if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                        log.LogError($"StackTrace: {e.StackTrace}");
                    }
                }

                if (ownerList.Count > 1)
                {
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
                            {"owners@odata.bind" , ownerList},
                            {"members@odata.bind" , memberList}
                        }
                    };

                    var result = await graphClient.Groups.Request().AddAsync(o365Group);
                    groupId = result.Id;
                    log.LogInformation($"Site and Office 365 {displayName} created successfully. And groupId: {groupId}");
                }
                else
                {
                    // if the "else" condition is true then it means only the creator account was added to the owner list
                    // so we want to stop the site creation process and update the status of the request
                    groupId = Convert.ToString(GroupCreationStatus.NoOwner);
                }
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                groupId = Convert.ToString(GroupCreationStatus.Unknown);
            }

            log.LogInformation($"CreateGroup processed a request. groupId: {groupId}");

            return groupId;
        }

        //public static async Task<bool> AddMembersToTeam(GraphServiceClient graphClient, ILogger log, string groupId, string teamId, string Members)
        //{
        //    log.LogInformation("AddMembersToTeam received a request.");

        //    try
        //    {
        //        foreach (string email in Members.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries))
        //        {
        //            var user  = await graphClient.Users[email].Request().GetAsync();
        //            var memberId = user.Id;

        //            log.LogInformation($"email: {email}");
        //            log.LogInformation($"memberId: {memberId}");

        //            var directoryObject = new DirectoryObject
        //            {
        //                Id = memberId
        //            };
        //            await graphClient.Groups[groupId].Members.References.Request().AddAsync(directoryObject);

        //            AadUserConversationMember mem = new AadUserConversationMember
        //            {
        //                Roles = new List<String>()
        //                {
        //                    "member"
        //                },
        //                AdditionalData = new Dictionary<string, object>()
        //                {
        //                    {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{memberId}')"}
        //                }
        //            };
        //            await graphClient.Teams[teamId].Members.Request().AddAsync(mem);
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        log.LogError($"Message: {e.Message}");
        //        if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
        //        log.LogError($"StackTrace: {e.StackTrace}");
        //    }

        //   log.LogInformation("AddMembersToTeam processed a request.");

        //    return true;
        //}

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

                teamsUrl = $@"https://teams.microsoft.com/#/l/team/{channelId}/conversations?groupId={teamId}&tenantId={tenantId}";
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

        public static async Task<bool> ApplyTemplate(ClientContext ctx, string queueName, string descriptionEn, string descriptionFr, string followingContentFeatureId, string teamsUrl, ExecutionContext functionContext, ILogger log)
        {
            log.LogInformation("ApplyTemplate received a request.");

            try
            {
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

                string PNP_TEMPLATE_FILE;

                if (queueName == "prob")
                    PNP_TEMPLATE_FILE = "template_prob.xml";
                else
                    PNP_TEMPLATE_FILE = "template_unclassified.xml";

                ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);
                log.LogInformation($"Successfully found template with ID '{template.Id}'");

                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
                {
                    ProgressDelegate = (message, progress, total) =>
                    {
                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
                    },
                    MessagesDelegate = (message, messageType) =>
                    {
                        switch (messageType)
                        {
                            case ProvisioningMessageType.Warning:
                                log.LogWarning(string.Format("{0} : {1}", message, web.Title));
                                break;
                            case ProvisioningMessageType.Error:
                                log.LogError(string.Format("{0} : {1}", message, web.Title));
                                break;
                            default:
                                log.LogInformation(string.Format("{0} - {1} : {2}", messageType, message, web.Title));
                                break;
                        }
                    },
                    SiteProvisionedDelegate = (title, url) =>
                    {
                        log.LogInformation(string.Format("Site provisioning complete: {0} - {1}", title, url));
                    },
                    ClearNavigation = true,
                    IgnoreDuplicateDataRowErrors = true
                };

                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");
                template.Connector = connector;

                template.Parameters.Add("DescriptionEn", descriptionEn);
                template.Parameters.Add("DescriptionFr", descriptionFr);
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

        public static async Task SiteToHubAssociation(ClientContext ctx, string hubSiteId, ILogger log)
        {
            log.LogInformation("SiteToHubAssociation received a request.");
            log.LogInformation("Site {siteurl} will be associated with hub {hubsiteID}", ctx.Url, hubSiteId);

            try
            {
                var pnpclient = PnPHttpClient.Instance.GetHttpClient(ctx);
                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, $"{ctx.Url}/_api/site/JoinHubSite('{hubSiteId}')")
                {
                    Content = null
                };
                request.Headers.Add("accept", "application/json;odata.metadata=none");
                request.Headers.Add("odata-version", "4.0");

                await PnPHttpClient.AuthenticateRequestAsync(request, ctx).ConfigureAwait(false);

                HttpResponseMessage response = await pnpclient.SendAsync(request, new CancellationToken());

                log.LogInformation("Site {siteurl} was successfully associated with hub {hubsiteID}", ctx.Url, hubSiteId);
            }
            catch (Exception e)
            {
                log.LogError($"Exception: {e.Message}");
                if (e.InnerException is not null)
                    log.LogError($"InnerException: {e.InnerException.Message}");
            }

            log.LogInformation("SiteToHubAssociation processed a request.");
        }

        public static async Task<IActionResult> AddToStatusQueue(string connectionString, string itemId, string status, ILogger log)
        {
            log.LogInformation("AddToStatusQueue received a request.");

            try
            {
                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            { "Id", itemId },
                            { "Status", status }
                        }
                    }
                };
                await Common.InsertMessageAsync(connectionString, "status", listItem, log);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("AddToStatusQueue processed a request.");

            return new OkResult();
        }

        public static async Task<bool> AddToTeamsLinkList(ROPCConfidentialTokenCredential tokenCredential, string siteId, string listId, string title, string teamsId, string teamsUrl, ILogger log)
        {
            log.LogInformation("AddToTeamsLinkList received a request.");

            bool result = true;

            try
            {
                var graphClient = new GraphServiceClient(tokenCredential);

                var listItem = new ListItem
                {
                    Fields = new FieldValueSet
                    {
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"Title", title},
                            {"TeamsID", teamsId},
                            {"Teamslink", teamsUrl}
                        }
                    }
                };

                await graphClient.Sites[siteId].Lists[listId].Items.Request().AddAsync(listItem);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null) log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
                result = false;
            }

            log.LogInformation("AddToTeamsLinkList processed a request.");

            return result;
        }
    }
}
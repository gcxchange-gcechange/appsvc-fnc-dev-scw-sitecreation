//using Azure.Identity;
//using Azure.Security.KeyVault.Secrets;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.Azure.WebJobs.Extensions.Http;
//using Microsoft.Azure.WebJobs;
//using Microsoft.Extensions.Logging;
//using Microsoft.SharePoint.Client;
//using PnP.Framework;
//using PnP.Framework.Provisioning.Connectors;
//using PnP.Framework.Provisioning.Model;
//using PnP.Framework.Provisioning.ObjectHandlers;
//using PnP.Framework.Provisioning.Providers.Xml;
//using System;
//using System.IO;
//using System.Reflection;
//using System.Security.Cryptography.X509Certificates;
//using System.Threading;
//using System.Threading.Tasks;
//using Microsoft.Extensions.Configuration;
//using Microsoft.Azure.Services.AppAuthentication;
//using Microsoft.Azure.KeyVault;
//using Microsoft.Azure.KeyVault.Models;
//using PnP.Core.Model.SharePoint;
//using PnP.Core;
//using PnP.Core.Services;

//using Microsoft.Extensions.DependencyInjection;
//using Microsoft.Extensions.Hosting;
//using Microsoft.Graph;
//using Microsoft.AspNetCore;
//using PnP.Core.Auth;
//using static Microsoft.AspNetCore.Hosting.Internal.HostingApplication;
//using CamlBuilder;
//using Microsoft.VisualBasic;
//using System.Linq;




//// No, you can update 1 place and, with time, it will all sync. It's just, I dont know what is faster.//
//// Change the group and wait for sharepoint and teams to sync.
//// Update sharepoint and wait for teams to sync... 
//// Update teams and wait for sharepoint to sync....

//// OK I have both of these update statements working. Do you have a sample of how I would do the "Update sharepoint" commanbd?
//// await graphClient.Groups[teamId].Request().UpdateAsync(group);
//// await graphClient.Teams[teamId].Request().UpdateAsync(team);

//// For sharepoint, it will be with the pnp framework and the app only
//// https://pnp.github.io/pnpcore/using-the-sdk/webs-intro.html


//// [9:36 AM] Postlethwaite, Oliver
//// OK, so after the site is created then Update Name gets called? Why would the name need to be updated after creation?
//// 
//// [9:38 AM] Lefebvre, Stéphanie
//// Yes, When we create the site, we create with the id as a name because we want the URL to be the id. We can't update the url after from graph, but we can update the name. 
//// 
//// [9:38 AM] Postlethwaite, Oliver
//// OK, that makes sense
//// 
//// [9:38 AM] Lefebvre, Stéphanie
//// But, if you can find a way to update the url, that would be best but I don't think it's possible
//// 
//// [9:46 AM] Postlethwaite, Oliver
//// I don't think so either but I'll check
//// 
//// [9:54 AM] Lefebvre, Stéphanie
//// Perfect, if you confirm we can't we can continue with this logic




//namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
//{
//    public static class ApplyTemplate
//    {

//        [FunctionName("ApplyTemplate")]
//        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.System, "get", "post", Route = null)] HttpRequest req, ILogger log, Microsoft.Azure.WebJobs.ExecutionContext functionContext)
//        {
//            log.LogInformation("ApplyTemplate processed a request.");

//            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

//            string aadApplicationId = config["clientId"];
//            string sharePointUrl = config["sharePointUrl"];
//            string tenantId = config["tenantId"];
//            string certificateName = config["certificateName"];
//            string keyVaultUrl = config["keyVaultUrl"];

//            // Get certificate from the key vault
//            X509Certificate2 mycert = await GetKeyVaultCertificateAsync(keyVaultUrl, certificateName, log);

//            string tenantName = "devgcx";

//            AuthenticationManager auth = new AuthenticationManager(aadApplicationId, mycert, $"{tenantName}.onmicrosoft.com");
//            ClientContext ctx = await auth.GetContextAsync(sharePointUrl);
//            ApplyProvisioningTemplate(ctx, log, functionContext, tenantId);

//            return new OkObjectResult("OK!");
//        }

//        internal static async Task<X509Certificate2> GetKeyVaultCertificateAsync(string keyVaultUrl, string name, ILogger log)
//        {
//            log.LogInformation("GetKeyVaultCertificateAsync processed a request.");

//            var serviceTokenProvider = new AzureServiceTokenProvider();
//            var keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(serviceTokenProvider.KeyVaultTokenCallback));

//            SecretBundle secret = await keyVaultClient.GetSecretAsync(keyVaultUrl, name);
//            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(secret.Value), string.Empty, X509KeyStorageFlags.MachineKeySet);
//            return certificate;

//            // If you receive the following error when running the Function;
//            // Microsoft.Azure.WebJobs.Host.FunctionInvocationException:
//            // Exception while executing function: NotificationFunctions.QueueOperation--->
//            // System.Security.Cryptography.CryptographicException:
//            // The system cannot find the file specified.at System.Security.Cryptography.NCryptNative.ImportKey(SafeNCryptProviderHandle provider, Byte[] keyBlob, String format) at System.Security.Cryptography.CngKey.Import(Byte[] keyBlob, CngKeyBlobFormat format, CngProvider provider)
//            //
//            // Please see https://stackoverflow.com/questions/31685278/create-a-self-signed-certificate-in-net-using-an-azure-web-application-asp-ne
//            // Add the following Application setting to the AF "WEBSITE_LOAD_USER_PROFILE = 1"
//        }

//        /// <summary>
//        /// This method will apply PNP template to a SharePoint site.
//        /// </summary>
//        /// <param name="ctx"></param>
//        /// <param name="log"></param>
//        /// <param name="functionContext"></param>
//        public static async void ApplyProvisioningTemplate(ClientContext ctx, ILogger log, Microsoft.Azure.WebJobs.ExecutionContext functionContext, string TENANT_ID)
//        {
//            try
//            {
//                //var web = await ctx.Web.GetAsync(p => p.Title);

//                ctx.RequestTimeout = Timeout.Infinite;
//                Web web = ctx.Web;
//                ctx.Load(web, w => w.Title);
//                ctx.ExecuteQuery();

//                log.LogInformation($"Successfully connected to site: {web.Title}");

//                DirectoryInfo dInfo;
//                var schemaDir = "";
//                string currentDirectory = functionContext.FunctionDirectory;


                


//                if (currentDirectory == null)
//                {
//                    log.LogInformation("NULL");
//                    string workingDirectory = Environment.CurrentDirectory;
//                    log.LogInformation($"workingDirectory: {workingDirectory}");
//                    currentDirectory = System.IO.Directory.GetParent(workingDirectory).Parent.Parent.FullName;
//                    dInfo = new DirectoryInfo(currentDirectory);
//                    schemaDir = dInfo + "\\GxDcCPS-SitesCreations-fnc\\bin\\Debug\\net461\\Templates\\GenericTemplate";
//                }
//                else
//                {

//                    log.LogInformation("NOT NULL");
//                    log.LogInformation($"currentDirectory: {currentDirectory}");    // C:\home\site\wwwroot\ApplyTemplate
//                    dInfo = new DirectoryInfo(currentDirectory);
//                    log.LogInformation($"dInfo1.Exists is {dInfo.Exists}");



//                    schemaDir = dInfo.Parent.FullName + "\\Templates\\GenericTemplate";
                    




//                }
                                
//                log.LogInformation($"schemaDir is {schemaDir}");                    // schemaDir is C:\home\site\wwwroot\Templates\GenericTemplate
//                DirectoryInfo dInfo2 = new DirectoryInfo(schemaDir);
                
//                log.LogInformation($"dInfo2.Exists is {dInfo2.Exists}");





//                XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");

//                //log.LogInformation($"sitesProvider.Uri is {sitesProvider.Uri}");
//                //log.LogInformation($"sitesProvider.Connector is {sitesProvider.Connector}");


//                log.LogInformation($"Getting the templates...");
//                var mytemplates = sitesProvider.GetTemplates();
//                log.LogInformation($"mytemplates.Count is {mytemplates.Count}");

//                for (int i = 0; i <= mytemplates.Count - 1; i++)
//                {
//                    log.LogInformation($"mytemplates[{i}].Description is {mytemplates[i].Description}");
//                }

//                string PNP_TEMPLATE_FILE = "template-name.xml";

//                log.LogInformation($"PNP_TEMPLATE_FILE is {PNP_TEMPLATE_FILE}");

//                //string FULL_PATH = $"{schemaDir}\\{PNP_TEMPLATE_FILE}";
//                //log.LogInformation($"FULL_PATH is {FULL_PATH}");
//                ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);

//                log.LogInformation($"Successfully found template with ID '{template.Id}'");

//                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
//                {
//                    ProgressDelegate = (message, progress, total) =>
//                    {
//                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
//                    }
//                };
//                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");

//                template.Connector = connector;

//                // string[] descriptions = description.Split('|');

//                //string ALL_USER_GROUP = ConfigurationManager.AppSettings["ALL_USER_GROUP"];
//                //string ASSIGNED_GROUP = ConfigurationManager.AppSettings["ASSIGNED_GROUP"];
//                //string HUB_URL = ConfigurationManager.AppSettings["HUB_URL"];
//                //string GCX_SUPPORT = ConfigurationManager.AppSettings["GCX_SUPPORT"];
//                //string GCX_SCA = ConfigurationManager.AppSettings["GCX_SCA"];


//                // Add site information
//                //template.Parameters.Add("descEN", descriptions[0]);
//                //template.Parameters.Add("descFR", descriptions[1]);
//                //template.Parameters.Add("TENANT_ID", TENANT_ID);
//                //template.Parameters.Add("ALL_USER_GROUP", ALL_USER_GROUP);
//                //template.Parameters.Add("ASSIGNED_GROUP", ASSIGNED_GROUP);
//                //template.Parameters.Add("HUB_URL", HUB_URL);
//                //template.Parameters.Add("GCX_SUPPORT", GCX_SUPPORT);
//                //template.Parameters.Add("GCX_SCA", GCX_SCA);


//                // Add user information
//                //template.Parameters.Add("UserOneId", ownerInfo[0]);
//                //template.Parameters.Add("UserOneName", ownerInfo[1]);
//                //template.Parameters.Add("UserOneMail", ownerInfo[2]);
//                //template.Parameters.Add("UserTwoId", ownerInfo[3]);
//                //template.Parameters.Add("UserTwoName", ownerInfo[4]);
//                //template.Parameters.Add("UserTwoMail", ownerInfo[5]);


//                // Unable to cast object of type 'PnP.Core.Model.SharePoint.Web' to type 'Microsoft.SharePoint.Client.Web'

                
//                log.LogInformation("ApplyProvisioningTemplate...");
//                web.ApplyProvisioningTemplate(template, ptai);

                


//                log.LogInformation("...worked!");


//                //web.WebTemplate.
//                //var team = await ctx.Team.GetAsync();
//                //var site = await ctx.Site.GetAsync();



                


//                log.LogInformation($"Site {web.Title} apply template successfully.");
//            }
//            catch (ReflectionTypeLoadException ex)
//            {
//                foreach (var item in ex.LoaderExceptions)
//                {
//                    log.LogInformation(item.Message);
//                }
//            }
//            catch (Exception e)
//            {
//                // 2022 - 11 - 07T19: 32:11Z[Information]   Message: The Provisioning Template URI template - name.xml is not valid.
//                log.LogInformation($"Message: {e.Message}");
//                if (e.InnerException is not null)
//                    log.LogInformation($"InnerException: {e.InnerException.Message}");

//            }
//        }

//        public static async void ApplyProvisioningTemplateOLD(ClientContext ctx, ILogger log, Microsoft.Azure.WebJobs.ExecutionContext functionContext, string TENANT_ID)
//        {
//            try
//            {
//                ctx.RequestTimeout = Timeout.Infinite;
//                Web web = ctx.Web;
//                ctx.Load(web, w => w.Title);
//                ctx.ExecuteQuery();

//                log.LogInformation($"Successfully connected to site: {web.Title}");

//                DirectoryInfo dInfo;
//                var schemaDir = "";
//                string currentDirectory = functionContext.FunctionDirectory;
//                if (currentDirectory == null)
//                {
//                    string workingDirectory = Environment.CurrentDirectory;
//                    currentDirectory = System.IO.Directory.GetParent(workingDirectory).Parent.Parent.FullName;
//                    dInfo = new DirectoryInfo(currentDirectory);
//                    schemaDir = dInfo + "\\GxDcCPS-SitesCreations-fnc\\bin\\Debug\\net461\\Templates\\GenericTemplate";
//                }
//                else
//                {
//                    dInfo = new DirectoryInfo(currentDirectory);
//                    schemaDir = dInfo.Parent.FullName + "\\Templates\\GenericTemplate";
//                }

//                log.LogInformation($"schemaDir is {schemaDir}");
//                XMLTemplateProvider sitesProvider = new XMLFileSystemTemplateProvider(schemaDir, "");
//                string PNP_TEMPLATE_FILE = "template-name.xml";
//                ProvisioningTemplate template = sitesProvider.GetTemplate(PNP_TEMPLATE_FILE);
//                log.LogInformation($"Successfully found template with ID '{template.Id}'");



//                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation
//                {
//                    ProgressDelegate = (message, progress, total) =>
//                    {
//                        log.LogInformation(string.Format("{0:00}/{1:00} - {2} : {3}", progress, total, message, web.Title));
//                    }
//                };
//                FileSystemConnector connector = new FileSystemConnector(schemaDir, "");

//                template.Connector = connector;

//                // string[] descriptions = description.Split('|');

//                //string ALL_USER_GROUP = ConfigurationManager.AppSettings["ALL_USER_GROUP"];
//                //string ASSIGNED_GROUP = ConfigurationManager.AppSettings["ASSIGNED_GROUP"];
//                //string HUB_URL = ConfigurationManager.AppSettings["HUB_URL"];
//                //string GCX_SUPPORT = ConfigurationManager.AppSettings["GCX_SUPPORT"];
//                //string GCX_SCA = ConfigurationManager.AppSettings["GCX_SCA"];


//                // Add site information
//                //template.Parameters.Add("descEN", descriptions[0]);
//                //template.Parameters.Add("descFR", descriptions[1]);
//                //template.Parameters.Add("TENANT_ID", TENANT_ID);
//                //template.Parameters.Add("ALL_USER_GROUP", ALL_USER_GROUP);
//                //template.Parameters.Add("ASSIGNED_GROUP", ASSIGNED_GROUP);
//                //template.Parameters.Add("HUB_URL", HUB_URL);
//                //template.Parameters.Add("GCX_SUPPORT", GCX_SUPPORT);
//                //template.Parameters.Add("GCX_SCA", GCX_SCA);


//                // Add user information
//                //template.Parameters.Add("UserOneId", ownerInfo[0]);
//                //template.Parameters.Add("UserOneName", ownerInfo[1]);
//                //template.Parameters.Add("UserOneMail", ownerInfo[2]);
//                //template.Parameters.Add("UserTwoId", ownerInfo[3]);
//                //template.Parameters.Add("UserTwoName", ownerInfo[4]);
//                //template.Parameters.Add("UserTwoMail", ownerInfo[5]);

//                web.ApplyProvisioningTemplate(template, ptai);

//                log.LogInformation($"Site {web.Title} apply template successfully.");
//            }
//            catch (ReflectionTypeLoadException ex)
//            {
//                foreach (var item in ex.LoaderExceptions)
//                {
//                    log.LogInformation(item.Message);
//                }



//            }
//        }













//    }
//}

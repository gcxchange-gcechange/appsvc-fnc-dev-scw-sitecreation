using Azure.Core;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.KeyVault.Models;
using Microsoft.Azure.KeyVault;
using Microsoft.Azure.Services.AppAuthentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            log.LogInformation("graphAuth processed a request.");


            IConfiguration config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables()
            .Build();

            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            var keyVaultUrl = config["keyVaultUrl"];
            var keyname = config["secretName"];
            var tenantid = config["tenantId"];
            var clientID = config["clientId"];

            SecretClientOptions optionsSecret = new SecretClientOptions()
            {
                Retry =
                {
                    Delay= TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                 }
            };

            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), optionsSecret);
            KeyVaultSecret secret = client.GetSecret(keyname);
            var clientSecret = secret.Value;
            // var clientSecret = config["clientSecret"];


            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantid, clientID, clientSecret, options);


            try
            {
                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
                return graphClient;
            }
            catch (Exception e)
            {
                log.LogInformation($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogInformation($"InnerException: {e.InnerException.Message}");
                return null;
            }

            



            


        }

        internal static async Task<X509Certificate2> GetKeyVaultCertificateAsync(string keyVaultUrl, string name, ILogger log)
        {
            log.LogInformation("GetKeyVaultCertificateAsync processed a request.");

            var serviceTokenProvider = new AzureServiceTokenProvider();
            var keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(serviceTokenProvider.KeyVaultTokenCallback));

            SecretBundle secret = await keyVaultClient.GetSecretAsync(keyVaultUrl, name);
            X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(secret.Value), string.Empty, X509KeyStorageFlags.MachineKeySet);
            return certificate;

            // If you receive the following error when running the Function;
            // Microsoft.Azure.WebJobs.Host.FunctionInvocationException:
            // Exception while executing function: NotificationFunctions.QueueOperation--->
            // System.Security.Cryptography.CryptographicException:
            // The system cannot find the file specified.at System.Security.Cryptography.NCryptNative.ImportKey(SafeNCryptProviderHandle provider, Byte[] keyBlob, String format) at System.Security.Cryptography.CngKey.Import(Byte[] keyBlob, CngKeyBlobFormat format, CngProvider provider)
            //
            // Please see https://stackoverflow.com/questions/31685278/create-a-self-signed-certificate-in-net-using-an-azure-web-application-asp-ne
            // Add the following Application setting to the AF "WEBSITE_LOAD_USER_PROFILE = 1"
        }

    }
}
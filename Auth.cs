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
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;

namespace appsvc_fnc_dev_scw_sitecreation_dotnet001
{
    internal class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            log.LogInformation("graphAuth processed a request.");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

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
            var clientSecretCredential = new ClientSecretCredential(tenantid, clientID, clientSecret, options);

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

       
        //internal static async Task<X509Certificate2> GetKeyVaultCertificateAsync(string keyVaultUrl, string name, ILogger log)
        //{
        //    log.LogInformation("GetKeyVaultCertificateAsync received a request.");

        //    var serviceTokenProvider = new AzureServiceTokenProvider();
        //    var keyVaultClient = new KeyVaultClient(new KeyVaultClient.AuthenticationCallback(serviceTokenProvider.KeyVaultTokenCallback));

        //    SecretBundle secret = await keyVaultClient.GetSecretAsync(keyVaultUrl, name);
        //    X509Certificate2 certificate = new X509Certificate2(Convert.FromBase64String(secret.Value), string.Empty, X509KeyStorageFlags.MachineKeySet);

        //    log.LogInformation("GetKeyVaultCertificateAsync processed a request.");

        //    return certificate;

        //    // If you receive the following error when running the Function;
        //    // Microsoft.Azure.WebJobs.Host.FunctionInvocationException:
        //    // Exception while executing function: NotificationFunctions.QueueOperation--->
        //    // System.Security.Cryptography.CryptographicException:
        //    // The system cannot find the file specified.at System.Security.Cryptography.NCryptNative.ImportKey(SafeNCryptProviderHandle provider, Byte[] keyBlob, String format) at System.Security.Cryptography.CngKey.Import(Byte[] keyBlob, CngKeyBlobFormat format, CngProvider provider)
        //    //
        //    // Please see https://stackoverflow.com/questions/31685278/create-a-self-signed-certificate-in-net-using-an-azure-web-application-asp-ne
        //    // Add the following Application setting to the AF "WEBSITE_LOAD_USER_PROFILE = 1"
        //}

        public class ROPCConfidentialTokenCredential : Azure.Core.TokenCredential
        {
            string _clientId;
            string _clientSecret;
            string _password;
            string _tenantId;
            string _tokenEndpoint;
            string _username;
            ILogger _log;

            //public ROPCConfidentialTokenCredential(ILogger log)
            //{
            //    IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            //    string keyVaultUrl = config["keyVaultUrl"];
            //    string secretName = config["secretName"];
            //    string secretNamePassword = config["secretNamePassword"];

            //    _clientId = config["clientId"];
            //    _tenantId = config["tenantId"];
            //    _username = config["user_name"];
            //    _log = log;
            //    _tokenEndpoint = "https://login.microsoftonline.com/" + _tenantId + "/oauth2/v2.0/token";

            //    SecretClientOptions options = new SecretClientOptions()
            //    {
            //        Retry =
            //    {
            //        Delay= TimeSpan.FromSeconds(2),
            //        MaxDelay = TimeSpan.FromSeconds(16),
            //        MaxRetries = 5,
            //        Mode = RetryMode.Exponential
            //     }
            //    };

            //    var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), options);

            //    KeyVaultSecret secret = client.GetSecret(secretName);
            //    _clientSecret = secret.Value;

            //    KeyVaultSecret password = client.GetSecret(secretNamePassword);
            //    _password = password.Value;
            //}

            public ROPCConfidentialTokenCredential(string userName, string userSecretName, ILogger log)
            {
                IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

                string keyVaultUrl = config["keyVaultUrl"];
                string secretName = config["secretName"];
                string secretNamePassword = userSecretName;

                _clientId = config["clientId"];
                _tenantId = config["tenantId"];
                _username = userName;
                _log = log;
                _tokenEndpoint = "https://login.microsoftonline.com/" + _tenantId + "/oauth2/v2.0/token";

                SecretClientOptions options = new SecretClientOptions()
                {
                    Retry =
                {
                    Delay= TimeSpan.FromSeconds(2),
                    MaxDelay = TimeSpan.FromSeconds(16),
                    MaxRetries = 5,
                    Mode = RetryMode.Exponential
                 }
                };

                var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), options);

                KeyVaultSecret secret = client.GetSecret(secretName);
                _clientSecret = secret.Value;

                KeyVaultSecret password = client.GetSecret(secretNamePassword);
                _password = password.Value;
            }


            public override AccessToken GetToken(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                HttpClient httpClient = new HttpClient();

                var Parameters = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                    new KeyValuePair<string, string>("username", _username),
                    new KeyValuePair<string, string>("password", _password),
                    new KeyValuePair<string, string>("grant_type", "password")
                };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
                {
                    Content = new FormUrlEncodedContent(Parameters)
                };

                var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
                dynamic responseJson = JsonConvert.DeserializeObject(response);
                var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
                return new AccessToken(responseJson.access_token.ToString(), expirationDate);
            }

            public override ValueTask<AccessToken> GetTokenAsync(TokenRequestContext requestContext, CancellationToken cancellationToken)
            {
                HttpClient httpClient = new HttpClient();

                var Parameters = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("scope", string.Join(" ", requestContext.Scopes)),
                    new KeyValuePair<string, string>("username", _username),
                    new KeyValuePair<string, string>("password", _password),
                    new KeyValuePair<string, string>("grant_type", "password")
                };

                HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, _tokenEndpoint)
                {
                    Content = new FormUrlEncodedContent(Parameters)
                };

                var response = httpClient.SendAsync(request).Result.Content.ReadAsStringAsync().Result;
                dynamic responseJson = JsonConvert.DeserializeObject(response);
                var expirationDate = DateTimeOffset.UtcNow.AddMinutes(60.0);
                return new ValueTask<AccessToken>(new AccessToken(responseJson.access_token.ToString(), expirationDate));
            }
        }
    }
}
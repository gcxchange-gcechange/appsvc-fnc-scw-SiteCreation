using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using PnP.Framework.Provisioning.Connectors;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace Test_App_only
{
    public static class ApplyTemplate
    {
        [FunctionName("ApplyTemplate")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.System, "get", "post", Route = null)] HttpRequest req, ILogger log, Microsoft.Azure.WebJobs.ExecutionContext functionContext)
        {
            string siteUrl = "https://devgcx.sharepoint.com/teams/stephtest";
            string clientId = "4bb150ca-b985-4d26-b306-ffde3119c570";
            string clientSecret = "1-h8Q~SYQY1azke6V2HRNTp1lv07nod-ICI6obWD";
            Web currentWeb;
            string aadApplicationId = "4bb150ca-b985-4d26-b306-ffde3119c570";
            string tenantName = "devgcx";
            string tenantID = "28d8f6f0-3824-448a-9247-b88592acc8b7";
            string sharePointUrl = $"https://devgcx.sharepoint.com/teams/MySampleTeam113";

            //ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(targetSiteUrl, appOnlyId, appOnlySecret);
            

            AuthenticationManager auth = new AuthenticationManager(aadApplicationId, GetCertificate(), $"{tenantName}.onmicrosoft.com");

            ClientContext ctx = await auth.GetContextAsync(sharePointUrl);
           
            ApplyProvisioningTemplate(ctx, log, functionContext, tenantID);

            return new OkObjectResult($"OK!");
        }//main


        private static X509Certificate2 GetCertificate()
        {
            string secretName = "app-only-test"; // Name of the certificate
            Uri keyVaultUri = new Uri($"https://dgcx-dev-keyvault-scw2.vault.azure.net/");

            var client = new SecretClient(keyVaultUri, new DefaultAzureCredential());
            KeyVaultSecret secret = client.GetSecret(secretName);

            return new X509Certificate2(Convert.FromBase64String(secret.Value), string.Empty, X509KeyStorageFlags.MachineKeySet);
        }

        /// <summary>
        /// This method will apply PNP template to a SharePoint site.
        /// </summary>
        /// <param name="ctx"></param>
        /// <param name="log"></param>
        /// <param name="functionContext"></param>
        public static async void ApplyProvisioningTemplate(ClientContext ctx, ILogger log, Microsoft.Azure.WebJobs.ExecutionContext functionContext, string TENANT_ID)
        {
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

                log.LogInformation($"schemaDir is {schemaDir}");
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

               // string[] descriptions = description.Split('|');

                //string ALL_USER_GROUP = ConfigurationManager.AppSettings["ALL_USER_GROUP"];
                //string ASSIGNED_GROUP = ConfigurationManager.AppSettings["ASSIGNED_GROUP"];
                //string HUB_URL = ConfigurationManager.AppSettings["HUB_URL"];
                //string GCX_SUPPORT = ConfigurationManager.AppSettings["GCX_SUPPORT"];
                //string GCX_SCA = ConfigurationManager.AppSettings["GCX_SCA"];


                // Add site information
                //template.Parameters.Add("descEN", descriptions[0]);
                //template.Parameters.Add("descFR", descriptions[1]);
                //template.Parameters.Add("TENANT_ID", TENANT_ID);
                //template.Parameters.Add("ALL_USER_GROUP", ALL_USER_GROUP);
                //template.Parameters.Add("ASSIGNED_GROUP", ASSIGNED_GROUP);
                //template.Parameters.Add("HUB_URL", HUB_URL);
                //template.Parameters.Add("GCX_SUPPORT", GCX_SUPPORT);
                //template.Parameters.Add("GCX_SCA", GCX_SCA);


                // Add user information
                //template.Parameters.Add("UserOneId", ownerInfo[0]);
                //template.Parameters.Add("UserOneName", ownerInfo[1]);
                //template.Parameters.Add("UserOneMail", ownerInfo[2]);
                //template.Parameters.Add("UserTwoId", ownerInfo[3]);
                //template.Parameters.Add("UserTwoName", ownerInfo[4]);
                //template.Parameters.Add("UserTwoMail", ownerInfo[5]);

                web.ApplyProvisioningTemplate(template, ptai);

                log.LogInformation($"Site {web.Title} apply template successfully.");
            }
            catch (ReflectionTypeLoadException ex)
            {
                foreach (var item in ex.LoaderExceptions)
                {
                    log.LogInformation(item.Message);
                }
            }
        }
    }//cs
    
}//n
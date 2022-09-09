using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using System.Collections.Generic;
using System.Linq;

namespace appsvc_fnc_scw_SiteCreation
{
    public static class CreateTeams
    {
        [FunctionName("CreateTeams")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            Auth auth = new Auth();
            var graphAPIAuth = auth.graphAuth(log);

            var callteams = createteams(graphAPIAuth, log).GetAwaiter().GetResult();
            var updateName = UpdateName(graphAPIAuth, callteams, log).GetAwaiter().GetResult();


            return new OkObjectResult($"OK!");
        }

        public static async Task<string> createteams(GraphServiceClient graphClient, ILogger log)
        {
            log.LogInformation("Call teams");
            var teamId = "";
            try
            {
                var team = new Team
                {
                    DisplayName = "Request ID",
                    Description = "My Sample Team’s Description2",
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
                            {"user@odata.bind", "https://graph.microsoft.com/v1.0/users('8ff97d6b-15ca-4042-a717-470aa8fcf6f9')"}
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
                    log.LogInformation(teamId);
                }

                return teamId;
            }
            catch (Exception ex)
            {
                log.LogInformation(ex.Message);
                return "not Good";
            }
        }

        public static async Task<string> UpdateName(GraphServiceClient graphClient, string teamsid, ILogger log)
        {

            var group = new Group
            {
                DisplayName = "Library Assist",
            };

            await graphClient.Groups[teamsid]
                .Request()
                .UpdateAsync(group);

            return "OK";
        }
    }
}

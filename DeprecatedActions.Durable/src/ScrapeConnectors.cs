using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using DeprecatedActions.Models;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace DeprecatedActions.Durable
{
    public static class ConnectorScrapeage
    {
        private readonly static Regex sConnectorUniqueName = new Regex(@"^\.\.\/(.*)\/$");

        [FunctionName(nameof(ScrapeConnectors))]
        public static async Task<IActionResult> ScrapeConnectors(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            string requestBody = new StreamReader(req.Body).ReadToEnd();
            var request = JsonConvert.DeserializeObject<ScrapeConnectorsRequest>(requestBody);

            log.LogInformation($"{nameof(ScrapeConnectors)} triggered at {DateTime.UtcNow:O}. " +
                               $"SelectedConnectors=${String.Join(",", request.SelectedConnectors)}");

            // Start the orchestrator with this request
            string instanceId = await starter.StartNewAsync("ScrapeConnectorsOrchestration", request);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            // Return a custom '202 Accepted' response
            return BuildAcceptedResponse(req, instanceId);
        }

        private static IActionResult BuildAcceptedResponse(HttpRequest req, string instanceId)
        {
            // Inform the client how long to wait in seconds before checking the status
            req.HttpContext.Response.Headers.Add("retry-after", "20");

            // Inform the client where to check again
            var location = string.Format("{0}://{1}/api/ScrapeConnectorsStatus/instance/{2}", req.Scheme, req.Host, instanceId);
            return new AcceptedResult(location, null);
        }

        /// Http Triggered Function which acts as a wrapper to get the status of a running Durable orchestration instance.
        /// We're using Anonymous Authorisation Level for demonstration purposes. You should use a more secure approach. 
        [FunctionName(nameof(ScrapeConnectorsStatus))]
        public static async Task<IActionResult> ScrapeConnectorsStatus(
            [HttpTrigger(AuthorizationLevel.Anonymous, methods: "get", Route = "ScrapeConnectorsStatus/instance/{instanceId}")] HttpRequest req,
            [DurableClient] IDurableOrchestrationClient orchestrationClient,
            string instanceId)
        {
            // Get the built-in status of the orchestration instance. This status is managed by the Durable Functions Extension. 
            var status = await orchestrationClient.GetStatusAsync(instanceId);
            if (status != null)
            {
                if (status.RuntimeStatus == OrchestrationRuntimeStatus.Running || status.RuntimeStatus == OrchestrationRuntimeStatus.Pending)
                {
                    return BuildAcceptedResponse(req, instanceId);
                }
                else if (status.RuntimeStatus == OrchestrationRuntimeStatus.Completed)
                {
                    return new OkObjectResult(status.Output);
                }
                else if (status.RuntimeStatus == OrchestrationRuntimeStatus.Failed)
                {
                    return new BadRequestObjectResult(status.Output);
                }
                throw new Exception($"Unexpected RuntimeStatus: {status.RuntimeStatus}");
            }

            // If status is null, then instance has not been found. Create and return an Http Response with status NotFound (404). 
            return new NotFoundObjectResult($"InstanceId {instanceId} is not found.");
        }

        //
        // Durable Function Orchestrator and Actions 
        //

        [FunctionName("ScrapeConnectorsOrchestration")]
        public static async Task<List<ConnectorInfo>> RunOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context, ILogger log)
        {
            // Scrape the connectors from the main page
            var connectors = await context.CallActivityAsync<List<ConnectorInfo>>(nameof(ScrapeConnectorsOrchestration_ScrapeConnectors), null);

            // Get the request parameters passed into the orchestrator
            var request = context.GetInput<ScrapeConnectorsRequest>();

            if (request.SelectedConnectors != null)
            {
                // Filter connectors found to just the selectedConnectors required
                connectors = connectors.Where(c => request.SelectedConnectors.Contains(c.UniqueName)).ToList();

                // Throw an exception if not all connectors are found, we could be more helpful with the error message here!!
                if (connectors.Count < request.SelectedConnectors.Count())
                {
                    throw new Exception("Not all connectors found");
                }
            }

            // Create a task for each connector to get the actions from their own docs page
            var tasks = new List<Task<ConnectorInfo>>();
            foreach (var connector in connectors)
            {
                tasks.Add(context.CallActivityAsync<ConnectorInfo>(nameof(ScrapeConnectorsOrchestration_GetConnectorInfo), connector));
            }

            // Wait for all the tasks to finish (Fan-In)
            await Task.WhenAll(tasks);

            // Put the results together and return them
            return tasks
                .Select(t => t.Result)
                .OrderBy(r => r.UniqueName)
                .ToList();
        }

        // Scrape Connection Reference page to get the names of all the connectors
        [FunctionName(nameof(ScrapeConnectorsOrchestration_ScrapeConnectors))]
        public static List<ConnectorInfo> ScrapeConnectorsOrchestration_ScrapeConnectors([ActivityTrigger] string name, ILogger log)
        {
            // https://www.jasongaylord.com/blog/2020/10/02/screen-scrape-dotnet-core-azure-function

            var url = @"https://docs.microsoft.com/en-us/connectors/connector-reference/";

            HtmlWeb web = new HtmlWeb();
            var htmlDoc = web.Load(url);

            var connectorNodes = htmlDoc.DocumentNode.SelectNodes("id('list-of-connectors')/following-sibling::table/tr/td/a");

            // We could etxract the
            //   icon
            //   availability in Logic Apps, Power automate and Power Apps
            //   whether connector is Preview
            //   whether connector is Premium
            // but don't do that.  We can probably extract from in the docs for each connector
            var relativeUrls = connectorNodes.Select(x => x.Attributes["href"].Value).ToList();

            log.LogInformation($"Found {relativeUrls.Count()} connectors");

            var connectorInfo = new List<ConnectorInfo>();

            foreach (var relativeUrl in relativeUrls)
            {
                var match = sConnectorUniqueName.Match(relativeUrl);
                if (!match.Success)
                {
                    log.LogInformation($"Could not determine uniqueName for relativeUrl {relativeUrl}");
                    continue;
                }

                var uniqueName = match.Groups[1].Value;

                // Note that there's a more direct URL at https://docs.microsoft.com/en-us/connectors/<uniqueName>/
                // but we'll continue to use the URL scrapped from the main connector reference which redirects
                var documentationUrl = $"{url}{relativeUrl}";

                // Populate basic information for the connector, we expand on this when the individual connector is scraped
                var c = new ConnectorInfo
                {
                    UniqueName = uniqueName,
                    DocumentationUrl = documentationUrl,
                    Actions = new List<ActionInfo>(),
                };
                connectorInfo.Add(c);
            }
            return connectorInfo;
        }

        // Get information about a single connector
        [FunctionName(nameof(ScrapeConnectorsOrchestration_GetConnectorInfo))]
        public async static Task<ConnectorInfo> ScrapeConnectorsOrchestration_GetConnectorInfo([ActivityTrigger] ConnectorInfo connector, ILogger log)
        {
            log.LogInformation($"Getting connector documentation for {connector.UniqueName}.");

            HtmlWeb web = new HtmlWeb();
            var htmlDoc = await web.LoadFromWebAsync(connector.DocumentationUrl);

            // Get the list of anchors from the initial table, from this we can get the:
            //   description
            //   if the action is deprecated
            // But we can't get the OperationId of the action
            var actionNodes = htmlDoc.DocumentNode.SelectNodes("id('actions')/following-sibling::table/tr");
            log.LogInformation($"Found {actionNodes?.Count() ?? 0} actions on connector {connector.UniqueName}");

            // Some connectors don't have any actions - abort early.
            if (actionNodes == null)
            {
                return connector;
            }

            // We have to convert to a List (not IEnumerable) because we're going to edit the contents of the list
            var actions = actionNodes.Select(x => new ActionInfo
            {
                Anchor = x.SelectSingleNode("td/a").GetAttributeValue("href", ""),
                Name = x.SelectSingleNode("td/a").InnerText,
                Description = x.SelectSingleNode("td/a/../following-sibling::td").InnerText.Trim(),
            }).ToList();

            // For all of the actions found find the operationId, and flag if it's deprecated
            foreach (var action in actions)
            {
                var anchorText = action.Anchor.Substring(1);

                var xpath = "id(\"" + anchorText + "\")/following-sibling::div/dl/dd";

                var operationId = htmlDoc.DocumentNode.SelectSingleNode(xpath)?.InnerText?.Trim() ?? "";

                action.OperationId = operationId;

                var lowerName = action.Name.ToLower();
                action.IsDeprecated = lowerName.Contains("[deprecated]") || lowerName.Contains("(deprecated)");
            }

            // Sort the actions within the connector
            connector.Actions = actions.OrderBy(a => a.OperationId).ToList();

            return connector;
        }
    }
}
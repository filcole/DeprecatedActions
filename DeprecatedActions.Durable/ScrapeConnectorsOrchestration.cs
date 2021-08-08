using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DeprecatedActions.Models;
using HtmlAgilityPack;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.DurableTask;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Octokit;

namespace DeprecatedActions.Durable
{
    public static class ScrapeConnectorsOrchestration
    {
        private readonly static Regex sConnectorUniqueName = new Regex(@"^\.\.\/(.*)\/$");

        [FunctionName("ScrapeConnectorsOrchestration")]
        public static async Task<List<ConnectorInfo>> RunOrchestrator(
            [OrchestrationTrigger] IDurableOrchestrationContext context, ILogger log)
        {
            // Scrape the connectors from the main page
            var connectors = await context.CallActivityAsync<List<ConnectorInfo>>(nameof(ScrapeConnectorsOrchestration_ScrapeConnectors), "bob");

            // Create a task for each connector to get the actions from their own docs page
            var tasks = new List<Task<ConnectorInfo>>();
            foreach (var connector in connectors.Take(10))
            {
                tasks.Add(context.CallActivityAsync<ConnectorInfo>(nameof(ScrapeConnectorsOrchestration_GetConnectorInfo), connector));
            }

            // Launch and wait for all the tasks to finish
            await Task.WhenAll(tasks);

            // Put the results together
            var connectors2 = tasks.Select(t => t.Result);

            // Update github
            await context.CallActivityAsync(nameof(ScrapeConnectorsOrchestration_UpdateGithub), connectors2);

            return connectors2.ToList();
        }

        public static List<ConnectorInfo> FilteredConnectors(List<ConnectorInfo> connectors, bool deprecatedOnly)
        {
            return connectors
                .Where(x => x.Actions.Any(a => a.IsDeprecated == deprecatedOnly))
                .Select(c => new ConnectorInfo
                {
                    UniqueName = c.UniqueName,
                    DocumentationUrl = c.DocumentationUrl,
                    Actions = c.Actions.Where(a => a.IsDeprecated == deprecatedOnly).ToList(),
                })
                .ToList();
        }

        [FunctionName(nameof(ScrapeConnectorsOrchestration_UpdateGithub))]
        public async static Task<bool> ScrapeConnectorsOrchestration_UpdateGithub([ActivityTrigger] List<ConnectorInfo> connectors, ILogger log)
        {
            // Source: https://laedit.net/2016/11/12/GitHub-commit-with-Octokit-net.html
            var gitHub = new GitHubClient(new ProductHeaderValue("MyCoolApp"));
            gitHub.Credentials = new Credentials("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");
            var owner = "filcole";
            var repo = "PowerAutomateConnectors";

            // *** Get the SHA fo the latest commit of the main branch
            var headMainRef = "heads/main";
            // Get reference of main branch
            var mainReference = await gitHub.Git.Reference.Get(owner, repo, headMainRef);
            // Get the latest commit of this branch
            var latestCommit = await gitHub.Git.Commit.Get(owner, repo, mainReference.Object.Sha);
            var currentTree = await gitHub.Git.Tree.GetRecursive(owner, repo, latestCommit.Tree.Sha);

            // *** Create the new blobs for files
            var allBlob = new NewBlob { Encoding = EncodingType.Utf8, Content = JsonConvert.SerializeObject(connectors, Formatting.Indented) };
            var allBlobRef = await gitHub.Git.Blob.Create(owner, repo, allBlob);

            var deprecatedConnectors = FilteredConnectors(connectors, true);
            var deprecatedBlob = new NewBlob { Encoding = EncodingType.Utf8, Content = JsonConvert.SerializeObject(deprecatedConnectors, Formatting.Indented) };
            var deprecatedBlobRef = await gitHub.Git.Blob.Create(owner, repo, deprecatedBlob);

            var nonDeprecatedConnectors = FilteredConnectors(connectors, false);
            var nonDeprecatedBlob = new NewBlob { Encoding = EncodingType.Utf8, Content = JsonConvert.SerializeObject(nonDeprecatedConnectors, Formatting.Indented) };
            var nonDeprecatedBlobRef = await gitHub.Git.Blob.Create(owner, repo, nonDeprecatedBlob);

            // ** Create a new tree with
            //      * the SHA of the tree of the latest commit as base
            //      * items based on the blob(s)

            // Create a new Tree

            // Don't copy from the current tree
            // https://github.com/octokit/octokit.net/issues/1610#issuecomment-305767094
            //  i.e. var nt = new NewTree { BaseTree = latestCommit.Tree.Sha };

            // Create a new tree from the current tree
            var nt = new NewTree();
            currentTree.Tree
                        .Where(x => x.Type != TreeType.Tree)
                        .Select(x => new NewTreeItem
                        {
                            Path = x.Path,
                            Mode = x.Mode,
                            Type = x.Type.Value,
                            Sha = x.Sha
                        })
                        .ToList()
                        .ForEach(x => nt.Tree.Add(x));

            // Add items based on blobs
            nt.Tree.Add(new NewTreeItem { Path = "All.json", Mode = "100644", Type = TreeType.Blob, Sha = allBlobRef.Sha });
            nt.Tree.Add(new NewTreeItem { Path = "Deprecated.json", Mode = "100644", Type = TreeType.Blob, Sha = deprecatedBlobRef.Sha });
            nt.Tree.Add(new NewTreeItem { Path = "NonDeprecated.json", Mode = "100644", Type = TreeType.Blob, Sha = nonDeprecatedBlobRef.Sha });

            var existingConnectorFiles = nt.Tree.Where(x => x.Path.StartsWith("connectors"));
            foreach (var toRemove in existingConnectorFiles.ToList())
            {
                nt.Tree.Remove(toRemove);
            }

            foreach (var connector in connectors)
            {
                log.LogInformation($"Creating blob for connector {connector.UniqueName}");
                var connectorBlob = new NewBlob { Encoding = EncodingType.Utf8, Content = JsonConvert.SerializeObject(connector, Formatting.Indented) };
                var connectorBlobRef = await gitHub.Git.Blob.Create(owner, repo, connectorBlob);

                nt.Tree.Add(new NewTreeItem
                {
                    Path = $"connectors/{connector.UniqueName}.json",
                    Mode = "100644",
                    Type = TreeType.Blob,
                    Sha = connectorBlobRef.Sha,
                });
            }

            //// Remove a file
            //var toRemove = nt.Tree.Where(x => x.Path.Equals("all.json")).FirstOrDefault();
            //if (toRemove != null)
            //{
            //    nt.Tree.Remove(toRemove);
            //}

            var newTree = await gitHub.Git.Tree.Create(owner, repo, nt);

            // ** Create the commit with the SHAs of the tree and the refernece of main branch

            // Create Commit
            var newCommit = new NewCommit($"Update as at {DateTime.UtcNow.ToString("O")}", newTree.Sha, mainReference.Object.Sha);
            var commit = await gitHub.Git.Commit.Create(owner, repo, newCommit);

            // ** Updte the reference of main branch with the SHA of the commit
            await gitHub.Git.Reference.Update(owner, repo, headMainRef, new ReferenceUpdate(commit.Sha));

            // FIXME: Put the above in a try/catch

            // No errors
            return true;
        }

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

            // Some connectors don't have any actions - abort early.
            if (actionNodes == null)
            {
                log.LogInformation($"No actions found on ${connector.UniqueName}");
                return connector;
            }

            log.LogInformation($"Found {actionNodes.Count()} actions");

            // We have to convert to a List (not IEnumerable) because we're going to edit the contents of the list
            var actions = actionNodes.Select(x => new ActionInfo
            {
                Anchor = x.SelectSingleNode("td/a").GetAttributeValue("href", ""),
                Name = x.SelectSingleNode("td/a").InnerText,
                Description = x.SelectSingleNode("td/a/../following-sibling::td").InnerText.Trim(),
            }).ToList();

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

        [FunctionName(nameof(ScrapeConnectorsOrchestration_HttpStart))]
        public static async Task<HttpResponseMessage> ScrapeConnectorsOrchestration_HttpStart(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post")] HttpRequestMessage req,
            [DurableClient] IDurableOrchestrationClient starter,
            ILogger log)
        {
            // Function input comes from the request content.
            string instanceId = await starter.StartNewAsync("ScrapeConnectorsOrchestration", null);

            log.LogInformation($"Started orchestration with ID = '{instanceId}'.");

            return starter.CreateCheckStatusResponse(req, instanceId);
        }
    }
}
using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Xml;
using HtmlAgilityPack;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace pgc
{
    public class ConnectorInfo
    {
        public string UniqueName { get; set; }
        public string DocumentationUrl { get; set; }
        public IEnumerable<ActionInfo> Actions { get; set; }
    }

    public class ActionInfo
    {
            public string Anchor { get; set; }
            public string Name{ get; set; }
            public string Description { get; set; }
            public string OperationId { get; set; }

            public bool IsDeprecated { get; set; }
    }

    public static class GetAllConnectors
    {
        //private readonly string Regex sConnectorUniqueName = new Regex();   // (@"^\.\.\/(.*)\/$");
        
        private readonly static Regex sConnectorUniqueName = new Regex(@"^\.\.\/(.*)\/$");

        private async static Task<List<ActionInfo>> GetConnectorInfo(string connectorReferenceUrl, ILogger log)
        {
            HtmlWeb web = new HtmlWeb();

            var htmlDoc = await web.LoadFromWebAsync(connectorReferenceUrl);

            // Get the list of anchors from the initial table, from this we can get the:
            //   description
            //   if the action is deprecated
            // But we can't get the OperationId of the action
            var actionNodes = htmlDoc.DocumentNode.SelectNodes("id('actions')/following-sibling::table/tr");

            // Some connectors don't have any actions
            if (actionNodes == null)
            {
                log.LogInformation($"No actions found on ${connectorReferenceUrl}");
                return null;
            }

            // We have to convert to a List (not IEnumerable) because we're going to edit the contents of the list
            var actions = actionNodes.Select(x => new ActionInfo
            {
                Anchor = x.SelectSingleNode("td/a").GetAttributeValue("href", ""),
                Name = x.SelectSingleNode("td/a").InnerText,
                Description = x.SelectSingleNode("td/a/../following-sibling::td").InnerText.Trim(),
            }).ToList();

            foreach (var action in actions) {

                #pragma warning disable IDE0057 // Use range operator
                var anchorText = action.Anchor.Substring(1);
                #pragma warning restore IDE0057 // Use range operator

                var xpath = "id(\"" + anchorText + "\")/following-sibling::div/dl/dd";

                var operationId = htmlDoc.DocumentNode.SelectSingleNode(xpath)?.InnerText?.Trim() ?? "";

                action.OperationId = operationId;

                var lowerName = action.Name.ToLower();
                action.IsDeprecated = lowerName.Contains("[deprecated]") || lowerName.Contains("(deprecated)");
            }
            return actions;
        }

        [FunctionName("GetAllConnectors")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            // Using https://html-agility-pack.net/
            //var url = @"http://html-agility-pack.net/";

            var url = @"https://docs.microsoft.com/en-us/connectors/connector-reference/";

            HtmlWeb web = new HtmlWeb();

            var htmlDoc = await web.LoadFromWebAsync(url);

            var connectorNodes= htmlDoc.DocumentNode.SelectNodes("id('list-of-connectors')/following-sibling::table/tr/td/a");

            // We could etxract the
            //   icon
            //   availability in Logic Apps, Power automate and Power Apps
            //   whether connector is Preview
            //   whether connector is Premium  
            // but don't do that.  We can probably extract from in the docs for each connector
            var relativeUrls = connectorNodes.Select(x => x.Attributes["href"].Value);

            log.LogInformation($"Found {relativeUrls.Count()} connectors");

            var connectorInfo = new List<ConnectorInfo>();
            var count = 0;

            foreach (var relativeUrl in relativeUrls)
            {
                var match = sConnectorUniqueName.Match(relativeUrl);
                if (!match.Success)
                {
                    log.LogWarning($"Could not determine uniqueName for relativeUrl {relativeUrl}");
                    continue;
                }
                
                var uniqueName = match.Groups[1].Value;

                // Note that there's a more direct URL at https://docs.microsoft.com/en-us/connectors/<uniqueName>/
                // but we'll continue to use the URL scrapped from the main connector reference which redirects
                var documentationUrl = $"{url}{relativeUrl}";

                log.LogInformation($"{++count}/{relativeUrls.Count()}: {documentationUrl}");

                connectorInfo.Add(new ConnectorInfo
                {
                    UniqueName = uniqueName,
                    DocumentationUrl = documentationUrl,
                    Actions = await GetConnectorInfo(documentationUrl, log),
                });
            }

            return new OkObjectResult(JsonConvert.SerializeObject(connectorInfo));
        }
    }
}

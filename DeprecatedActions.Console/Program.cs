using System;
using System.IO;
using System.Threading.Tasks;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.Azure.WebJobs;
//using Microsoft.Azure.WebJobs.Extensions.Http;
//using Microsoft.AspNetCore.Http;
//using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Xml;
using HtmlAgilityPack;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
//using System.Diagnostics.Debug;
using DeprecatedActions.Models;

namespace DeprecatedActions.Prog
{
    class Program
    {
        private readonly static string BASEDIR = @"c:\dev\Repos\Connectors.Rip\";

        private readonly static Regex sConnectorUniqueName = new Regex(@"^\.\.\/(.*)\/$");

        static async Task Main(string[] args)
        {
            var url = @"https://docs.microsoft.com/en-us/connectors/connector-reference/";

            HtmlWeb web = new HtmlWeb();

            var htmlDoc = await web.LoadFromWebAsync(url);

            var connectorNodes = htmlDoc.DocumentNode.SelectNodes("id('list-of-connectors')/following-sibling::table/tr/td/a");

            // We could etxract the
            //   icon
            //   availability in Logic Apps, Power automate and Power Apps
            //   whether connector is Preview
            //   whether connector is Premium  
            // but don't do that.  We can probably extract from in the docs for each connector
            var relativeUrls = connectorNodes.Select(x => x.Attributes["href"].Value).ToList();

            Console.WriteLine($"Found {relativeUrls.Count()} connectors");

            var connectorInfo = new List<ConnectorInfo>();
            var count = 0;

            relativeUrls.Sort();

            foreach (var relativeUrl in relativeUrls.Take(10))
            {
                var match = sConnectorUniqueName.Match(relativeUrl);
                if (!match.Success)
                {
                    Console.WriteLine($"Could not determine uniqueName for relativeUrl {relativeUrl}");
                    continue;
                }

                var uniqueName = match.Groups[1].Value;

                // Note that there's a more direct URL at https://docs.microsoft.com/en-us/connectors/<uniqueName>/
                // but we'll continue to use the URL scrapped from the main connector reference which redirects
                var documentationUrl = $"{url}{relativeUrl}";

                Console.WriteLine($"{++count}/{relativeUrls.Count()}: {documentationUrl}");

                var c = new ConnectorInfo
                {
                    UniqueName = uniqueName,
                    DocumentationUrl = documentationUrl,
                    Actions = await GetConnectorInfo(documentationUrl),
                };
                connectorInfo.Add(c);

                await File.WriteAllTextAsync(Path.Join(BASEDIR, "connectors", $"{uniqueName.Trim()}.json"), JsonConvert.SerializeObject(c, Newtonsoft.Json.Formatting.Indented));
            }

            // Giant list of all connectors
            await File.WriteAllTextAsync(Path.Join(BASEDIR, "all.json"), JsonConvert.SerializeObject(connectorInfo, Newtonsoft.Json.Formatting.Indented));

            await File.WriteAllTextAsync(Path.Join(BASEDIR, "deprecated.json"), JsonConvert.SerializeObject(
                connectorInfo
                    .Where(x => x.Actions.Any(a => a.IsDeprecated))
                    .Select(c => new ConnectorInfo
                    {
                            UniqueName = c.UniqueName,
                            DocumentationUrl = c.DocumentationUrl,
                            Actions = c.Actions.Where(a => a.IsDeprecated).ToList(),
                    }),
                Newtonsoft.Json.Formatting.Indented
            ));

            await File.WriteAllTextAsync(Path.Join(BASEDIR, "current.json"), JsonConvert.SerializeObject(
                connectorInfo
                    .Where(x => x.Actions.Any(a => !a.IsDeprecated))
                    .Select(c => new ConnectorInfo
                    {
                        UniqueName = c.UniqueName,
                        DocumentationUrl = c.DocumentationUrl,
                        Actions = c.Actions.Where(a => !a.IsDeprecated).ToList(),
                    }), Newtonsoft.Json.Formatting.Indented
            ));

        }

        private async static Task<List<ActionInfo>> GetConnectorInfo(string connectorReferenceUrl)
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
                Console.WriteLine($"No actions found on ${connectorReferenceUrl}");
                return new List<ActionInfo>();
            }

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

            return actions.OrderBy(a => a.OperationId).ToList();
        }
    }
}

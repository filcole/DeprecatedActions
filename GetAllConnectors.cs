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
        private readonly static Regex sWhitespace = new Regex(@"\s+");
        private readonly static Regex sSingleQuote = new Regex(@"\'");

        private static IEnumerable<ActionInfo> GetConnectorInfo(string connectorReferenceUrl)
        {
            HtmlWeb web = new HtmlWeb();

            var htmlDoc = web.Load(connectorReferenceUrl);

            // Get the list of anchors from the initial table, from this we can get the:
            //   description
            //   if the action is deprecated
            // But we can't get the OperationId of the action
            var actionNodes = htmlDoc.DocumentNode.SelectNodes("id('actions')/following-sibling::table/tr");

            // Some connectors don't have any actions
            if (actionNodes == null)
            {
                return null;
            }

            var actions = actionNodes.Select(x => new ActionInfo
            {
                Anchor = x.SelectSingleNode("td/a").GetAttributeValue("href", ""),
                Name = x.SelectSingleNode("td/a").InnerText,
                Description = sWhitespace.Replace(x.SelectSingleNode("td/a/../following-sibling::td").InnerText, ""),
            });
                

            // Find the action inforation by searching for the anchor
            //   extract out the Operation Id
            foreach (var action in actions) {

                //DELETEME var anchorText = sSingleQuote.Replace(action.Anchor.Substring(1), @"\'\'");
                // Remove the # from the beginning of the anchor
                var anchorText = action.Anchor.Substring(1);

                var xpath = "id(\"" + anchorText + "\")/following-sibling::div/dl/dd";

                var operationId = htmlDoc.DocumentNode.SelectSingleNode(xpath)?.InnerText ?? "";

                action.OperationId = sWhitespace.Replace(operationId, "");

                var lowerDescription = action.Name.ToLower();
                action.IsDeprecated = lowerDescription.Contains("[deprecated]") || lowerDescription.Contains("(deprecated)");
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

            var htmlDoc = web.Load(url);

            var connectorNodes= htmlDoc.DocumentNode.SelectNodes("id('list-of-connectors')/following-sibling::table/tr/td/a");

            // We could etxract the
            //   icon
            //   availability in Logic Apps, Power automate and Power Apps
            //   whether connector is Preview
            //   whether connector is Premium  
            // but don't do that.  We can probably extract from in the docs for each connector
            var connectorPages = connectorNodes.Select(x => x.Attributes["href"].Value);

            log.LogInformation($"Found {connectorPages.Count()} connectors");

            var connectorInfo = new List<ConnectorInfo>();
            var count = 0;

            foreach (var connectorPage in connectorPages)
            {
                var match = sConnectorUniqueName.Match(connectorPage);
                

                var documentationUrl = $"{url}{connectorPage}";

                log.LogInformation($"{++count}/{connectorPages.Count()}: {documentationUrl}");

                connectorInfo.Add(new ConnectorInfo
                {
                    UniqueName = connectorPage,
                    DocumentationUrl = documentationUrl,
                    Actions = GetConnectorInfo(documentationUrl),
                });
            }

             

            //string name = req.Query["name"];

            //string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            //dynamic data = JsonConvert.DeserializeObject(requestBody);
            //name = name ?? data?.name;

//            string responseMessage = string.IsNullOrEmpty(name)
//                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
//                : $"Hello, {name}. This HTTP triggered function executed successfully.";

            return new OkObjectResult(JsonConvert.SerializeObject(connectorInfo));
        }
    }
}

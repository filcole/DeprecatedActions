using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DeprecatedActions.Models
{

    public class ConnectorInfo
    {
        public string UniqueName { get; set; }
        public string DocumentationUrl { get; set; }
        public List<ActionInfo> Actions { get; set; }
    }

    public class ActionInfo
    {
        public string OperationId { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Anchor { get; set; }
        public bool IsDeprecated { get; set; }
    }
}


using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Text;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Internal.Writers
{
    internal class PluginPackageRelationship : IPluginPackageRelationship
    {
        public string RelationshipId {get; set;}

        public string RelationshipType { get; set; }

        public string Target { get; set; }

        public TargetMode TargetMode { get; set; }
    }
}

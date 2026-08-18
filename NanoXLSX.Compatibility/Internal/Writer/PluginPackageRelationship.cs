/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using System.IO.Packaging;

namespace NanoXLSX.Internal.Writer
{
    internal class PluginPackageRelationship : IPluginPackageRelationship
    {
        public string RelationshipId { get; set; }

        public string RelationshipType { get; set; }

        public string Target { get; set; }

        public TargetMode TargetMode { get; set; }
    }
}

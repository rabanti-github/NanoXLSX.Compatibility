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
    /// <summary>
    /// Class representing the interface implementation of <see cref="IPluginPackageRelationship"/>
    /// </summary>
    internal class PluginPackageRelationship : IPluginPackageRelationship
    {
        /// <summary>
        /// Relationship ID (rId)
        /// </summary>
        public string RelationshipId { get; set; }

        /// <summary>
        /// Relationship type URI (definition)
        /// </summary>
        public string RelationshipType { get; set; }

        /// <summary>
        /// Target URI
        /// </summary>
        public string Target { get; set; }

        /// <summary>
        /// Target mode of the relationship URI
        /// </summary>
        public TargetMode TargetMode { get; set; }
    }
}

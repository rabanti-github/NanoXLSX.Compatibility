/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Internal.Writer
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_PACKAGE_REGISTRY", QueueUUID = PlugInUUID.WriterPackageRegistryQueue, PlugInOrder = 20000)]
    internal class ExternalLinkPackageRegistry : IPluginPackageRegistry
    {
        #region constants
        private const string packagePartPath = "xl/externalLinks/";
        private const string contentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        private const string relationshipType = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        private const string relationshipPathType = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        #endregion
        #region privateFields
        private List<ExternalLink> externalLinks;

        private List<int> orderNumbers = new List<int>();

        private List<string> packagePartPaths = new List<string>();

        private List<string> packagePartFileNames = new List<string>();

        private List<string> contentTypes = new List<string>();

        private List<string> relationshipTypes = new List<string>();

        private List<bool> arePackagePartsRoot = new List<bool>();

        private List<string> uniquePackagePartIndices = new List<string>();

        private List<List<PluginPackageRelationship>> packagePartRelationships = new List<List<PluginPackageRelationship>>();
        #endregion
        #region properties
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// List of order numbers to be used for package registry
        /// </summary>
        public IReadOnlyList<int> OrderNumbers => orderNumbers;
        /// <summary>
        /// List of package part paths to be used for package registry
        /// </summary>
        public IReadOnlyList<string> PackagePartPaths => packagePartPaths;
        /// <summary>
        /// List of package part file names to be used for package registry
        /// </summary>
        public IReadOnlyList<string> PackagePartFileNames => packagePartFileNames;
        /// <summary>
        /// List of content types (URI) to be used for package registry
        /// </summary>
        public IReadOnlyList<string> ContentTypes => contentTypes;

        /// <summary>
        /// List of relationship types (URI) to be used for package registry
        /// </summary>
        public IReadOnlyList<string> RelationshipTypes => relationshipTypes;
        /// <summary>
        /// List of booleans, indicating whether the package parts are assigned to root or not, to be used for package registry
        /// </summary>
        public IReadOnlyList<bool> ArePackagePartsRoot => arePackagePartsRoot;
        /// <summary>
        /// List of unique indices to identify the package parts, to be used for package registry
        /// </summary>
        public IReadOnlyList<string> UniquePackagePartIndices => uniquePackagePartIndices;
        /// <summary>
        /// List of package part relationships of type <see cref="PluginPackageRelationship"/>, to be used for package registry
        /// </summary>
        public IReadOnlyList<IReadOnlyList<IPluginPackageRelationship>> PackagePartRelationships => packagePartRelationships;
        #endregion
        #region methods
        /// <summary>
        /// Initializing method (interface implementation)
        /// </summary>
        /// <param name="baseWriter">Base writer</param>
        public void Init(IBaseWriter baseWriter)
        {
            this.Workbook = baseWriter.Workbook;
        }

        /// <summary>
        /// Main execution method of the processor (interface implementation)
        /// </summary>
        public void Execute()
        {
            int orderNumber = Workbook.AuxiliaryData.GetData<int>(PlugInUUID.WriterPackageRegistryQueue, PlugInUUID.LastPackageOrderNumber);
            List<ExternalLink> storedExternalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
            externalLinks = storedExternalLinks == null ? new List<ExternalLink>() : storedExternalLinks.OfType<ExternalLink>().ToList();

            for (int i = 0; i < externalLinks.Count; i++)
            {
                orderNumber++;
                string name = "externalLink" + ParserUtils.ToString(i + 1) + ".xml";
                string uniqueIndex = CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + ParserUtils.ToString(i);
                orderNumbers.Add(orderNumber);
                packagePartPaths.Add(packagePartPath);
                packagePartFileNames.Add(name);
                contentTypes.Add(contentType);
                relationshipTypes.Add(relationshipType);
                arePackagePartsRoot.Add(false);
                uniquePackagePartIndices.Add(uniqueIndex);
                // package relationships
                ExternalLink externalLink = externalLinks[i];
                List<PluginPackageRelationship> relationshipList = new List<PluginPackageRelationship>();
                foreach (ExternalLinkUriRelationship rawRelationship in externalLink.GetUriRelationships())
                {
                    PluginPackageRelationship relationship = new PluginPackageRelationship();
                    relationship.RelationshipId = rawRelationship.Id;
                    relationship.Target = rawRelationship.SerializedTarget;
                    relationship.TargetMode = System.IO.Packaging.TargetMode.External;
                    relationship.RelationshipType = relationshipPathType;
                    relationshipList.Add(relationship);
                }
                if (relationshipList.Count > 0)
                {
                    packagePartRelationships.Add(relationshipList);
                }
            }
        }
        #endregion
    }
}

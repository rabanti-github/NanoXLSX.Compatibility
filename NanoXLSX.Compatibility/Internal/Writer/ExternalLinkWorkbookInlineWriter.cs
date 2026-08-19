/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;

namespace NanoXLSX.Internal.Writer
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_INLINE_WORKBOOK_WRITER", QueueUUID = PlugInUUID.WorkbookInlineWriter, PlugInOrder = 100000)]
    internal class ExternalLinkWorkbookInlineWriter : IPluginInlineWriter
    {
        #region properties
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Write context
        /// </summary>
        public IWriteContext WriteContext { get; set; } // NoOp
        /// <summary>
        /// Root element of the parent writer
        /// </summary>
        public XmlElement RootElement { get; set; }
        /// <summary>
        /// Current XML element
        /// </summary>
        public XmlElement XmlElement { get; } // NoOp
        #endregion
        #region methods
        /// <summary>
        /// Initializing method (interface implementation)
        /// </summary>
        /// <param name="rootElement">Root element of the parent writer</param>
        /// <param name="workbook">Current workbook</param>
        /// <param name="index">Index applied in <see cref="Execute"/></param>
        public void Init(ref XmlElement rootElement, Workbook workbook, int? index = null)
        {
            Workbook = workbook;
            RootElement = rootElement;
        }

        /// <summary>
        /// Main execution method of the processor (interface implementation)
        /// </summary>
        public void Execute()
        {
            XmlElement externalReferences = GetExternalReferences();
            if (externalReferences == null)
            {
                return; // Nothing to do
            }
            // XSD requires externalReferences to be after (mandatory) sheets and before (already existing) defined names 
            RootElement.AddChildElementAfter(externalReferences, "functionGroups", "sheets"); // Add <externalReferences> to <workbook>
            ReplaceDefinedNames();
            // TODO add further workbook-related processing, if external links are somewhere else too
        }

        /// <summary>
        ///  Method to replace defined name expressions (formulas) with human-readable external link markers into the OOXML compliant form (e.g. [1])
        /// </summary>
        private void ReplaceDefinedNames()
        {
            Dictionary<int, ExternalLinkResolution> externalLinks = Workbook.AuxiliaryData.GetData<Dictionary<int, ExternalLinkResolution>>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY);
            if (!Workbook.Features.ContainsDefinedNames || externalLinks == null || externalLinks.Count == 0)
            {
                return; // No defined names or external links to process
            }
            IEnumerable<XmlElement> definedNames = RootElement.FindChildElementsByNameAndAttribute("definedName", "name");
            int index = 0;
            foreach (XmlElement definedName in definedNames)
            {
                if (externalLinks.TryGetValue(index, out ExternalLinkResolution value))
                {
                    definedName.InnerValue = value.Expression;
                }
                index++;
            }
        }
        /// <summary>
        /// Method to get the XML element of external references, to be written into the workbook XML 
        /// </summary>
        /// <returns>XML element</returns>
        /// <exception cref="IOException">Thrown if the expected ExternalLink package part is not avalilable</exception>
        private XmlElement GetExternalReferences()
        {
            List<ExternalLink> externalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
            if (externalLinks.Count == 0)
            {
                return null;
            }

            XmlElement externalReferences = XmlElement.CreateElement("externalReferences");
            for (int i = 0; i < externalLinks.Count; i++)
            {
                string uniquePackagePartIndex = CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + ParserUtils.ToString(i);
                string relationshipId = Workbook.AuxiliaryData.GetData<string>(PlugInUUID.WriterPackageRegistryQueue, PlugInUUID.PackagePartRelationshipId, uniquePackagePartIndex);
                if (string.IsNullOrWhiteSpace(relationshipId))
                {
                    throw new IOException("The workbook relationship ID for external link package part '" + uniquePackagePartIndex + "' is not available.");
                }
                externalReferences.AddChildElementWithAttribute("externalReference", "id", relationshipId, "", "r");
            }
            return externalReferences;
        }
        #endregion
    }
}

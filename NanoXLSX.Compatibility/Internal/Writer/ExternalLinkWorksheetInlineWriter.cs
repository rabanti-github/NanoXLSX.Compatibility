/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Internal.Writer
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_INLINE_WORKSHEET_WRITER", QueueUUID = PlugInUUID.WorksheetInlineWriter, PlugInOrder = 100000)]
    internal class ExternalLinkWorksheetInlineWriter : IPluginInlineWriter
    {
        #region privateFields
        private Worksheet currentWorksheet;
        #endregion
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
            // Note: The worksheet writer or will pass the 1-based sheetID, not the 0-based index
            currentWorksheet = Workbook.Worksheets.First(w => w.SheetID == index); // This should never fail
        }

        /// <summary>
        /// Main execution method of the processor (interface implementation)
        /// </summary>
        public void Execute()
        {
            if (!currentWorksheet.Features.ContainsExternalLinks)
            {
                return; // Nothing to do
            }
            ReplaceFormulas();
            // TODO add further worksheet-related processing, if external links are somewhere else too
        }

        /// <summary>
        /// Method to replace formulas with human-readable external link markers into the OOXML compliant form (e.g. [1])
        /// </summary>
        private void ReplaceFormulas()
        {
            Dictionary<int, Dictionary<string, ExternalLinkResolution>> externalLinks =
                Workbook.AuxiliaryData.GetData<Dictionary<int, Dictionary<string, ExternalLinkResolution>>>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY);
            if (externalLinks == null || externalLinks.Count == 0 || !externalLinks.TryGetValue(currentWorksheet.SheetID, out Dictionary<string, ExternalLinkResolution> externalLink))
            {
                return; // No formulas or external links to process
            }
            foreach (KeyValuePair<string, ExternalLinkResolution> extLinkFormula in externalLink)
            {
                IEnumerable<XmlElement> cells = RootElement.FindChildElementsByNameAndAttribute("c", "r", extLinkFormula.Key);
                // This should never be ambiguous or non-existing
                XmlElement cell = cells.FirstOrDefault();
                if (cell != null)
                {
                    IEnumerable<XmlElement> formulas = cell.FindChildElementsByName("f");
                    // This should never be ambiguous or non-existing
                    XmlElement formula = formulas.FirstOrDefault();
                    if (formula != null)
                    {
                        formula.InnerValue = extLinkFormula.Value.Expression;
                    }
                }
            }
        }
        #endregion

    }
}

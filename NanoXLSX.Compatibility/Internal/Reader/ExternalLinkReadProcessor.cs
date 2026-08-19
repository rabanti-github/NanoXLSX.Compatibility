/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Internal.Reader
{
    /// <summary>
    /// Class implementing a finalizing processor for external links after reading.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_READ_PROCESSOR", QueueUUID = PlugInUUID.FinalizingInlineProcessor, PlugInOrder = 20000)]
    internal class ExternalLinkReadProcessor : IPluginInlineReadProcessor
    {
        public Workbook Workbook { get; set; }

        public void Init(Workbook workbook, IOptions readerOptions, int? index = null)
        {
            this.Workbook = workbook;
        }

        public void Execute()
        {
            List<ExternalLink> externalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
            List<string> externalReferenceRids = Workbook.AuxiliaryData.GetData<List<string>>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY);
            if (externalLinks == null || externalReferenceRids == null)
            {
                return; // No (valid) external links in workbook
            }
            Dictionary<string, HashSet<string>> definedNameMap = MapRidsToDefinedNames(externalReferenceRids, Workbook.GetDefinedNames());
            Dictionary<string, ExternalLink> externalReferences = new Dictionary<string, ExternalLink>();
            Dictionary<string, DefinedName> replacementMap = new Dictionary<string, DefinedName>();
            PrepareUpdateDefinedNames(externalLinks, externalReferenceRids, externalReferences, replacementMap);
            if (Workbook.Features.ContainsWorksheetFormulas && externalReferences.Count != 0 && replacementMap.Count > 0)
            {
                UpdateCellFormulas(externalReferences, replacementMap);
            }
            UpdateDefinedNames(replacementMap);

        }

        private void PrepareUpdateDefinedNames(List<ExternalLink> externalLinks, List<string> externalReferenceRids, Dictionary<string, ExternalLink> externalReferences, Dictionary<string, DefinedName> replacementMap)
        {
            for (int i = 0; i < externalReferenceRids.Count; i++)
            {
                ExternalLink link = externalLinks.Where(l => externalReferenceRids[i].Equals(l.WorkbookRId, StringComparison.OrdinalIgnoreCase)).First();
                if (link == null)
                {
                    throw new IOException("Mismatch of read external links and references in workbooks detected. External references cannot be resolved.");
                }
                string rId = "[" + ParserUtils.ToString(i) + "]";
                externalReferences[rId] = link;
            }
            IReadOnlyList<DefinedName> definedNames = Workbook.GetDefinedNames();
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName defiedName = definedNames[i];
                if (defiedName.Features.ContainsExternalLinks)
                {
                    if (ExternalLinkFormulaUtils.DetectExternalLinkId(defiedName.TextValue))
                    {
                        string replacedExpression = ExternalLinkFormulaUtils.ReplaceExternalLinkId(defiedName.TextValue, externalReferences);
                        defiedName.ReplaceExpression(replacedExpression);
                        string id = GetDefinedNameId(defiedName);
                        replacementMap[id] = defiedName;
                    }
                }
            }
        }

        private void UpdateCellFormulas(Dictionary<string, ExternalLink> externalReferences, Dictionary<string, DefinedName> replacementMap)
        {
            foreach (Worksheet worksheet in Workbook.Worksheets)
            {
                foreach (KeyValuePair<string, Cell> cell in worksheet.Cells)
                {
                    if (cell.Value.DataType == Cell.CellType.Formula)
                    {
                        if (cell.Value.Formula == null && cell.Value.Value != null)
                        {
                            string originalValue = cell.Value.Value?.ToString() ?? string.Empty;
                            if (ExternalLinkFormulaUtils.DetectExternalLinkId(originalValue))
                            {
                                string replacement = ExternalLinkFormulaUtils.ReplaceExternalLinkId(originalValue, externalReferences);
                                FormulaData formulaData = new FormulaData(replacement); // No cached value available
                                formulaData.HasExternalReferences = true;
                                cell.Value.Formula = formulaData;
                                cell.Value.Value = replacement;

                            }
                        }
                        else if (cell.Value.Formula != null)
                        {
                            string originalValue = cell.Value.Value?.ToString() ?? string.Empty;
                            if (ExternalLinkFormulaUtils.DetectExternalLinkId(originalValue))
                            {
                                string replacement = ExternalLinkFormulaUtils.ReplaceExternalLinkId(originalValue, externalReferences);
                                if (object.Equals(cell.Value.Formula.Expression, cell.Value.Value))
                                {
                                    cell.Value.Value = replacement;
                                }
                                cell.Value.Formula.Expression = replacement;
                                cell.Value.Formula.HasExternalReferences = true;
                            }
                            if (replacementMap.Count > 0 && cell.Value.Formula.DefinedNameReference != null)
                            {
                                string id = GetDefinedNameId(cell.Value.Formula.DefinedNameReference);
                                if (replacementMap.TryGetValue(id, out DefinedName newValue))
                                {
                                    cell.Value.Formula.DefinedNameReference = newValue;
                                }
                            }
                        }
                    }
                }
            }
        }

        private void UpdateDefinedNames(Dictionary<string, DefinedName> replacementMap)
        {
            if (replacementMap.Count == 0)
            {
                //return;
            }
            for (int i = Workbook.GetDefinedNames().Count - 1; i > 0; i--)
            {
                string id = GetDefinedNameId(Workbook.GetDefinedNames()[i]);
                if (replacementMap.TryGetValue(id, out DefinedName newValue))
                {
                    Workbook.RemoveDefinedName(newValue.Name, false, newValue.LocalSheet); // Prevent removing formula references
                    Workbook.AddDefinedName(newValue);
                }
            }
        }

        private static string GetDefinedNameId(DefinedName defiedName)
        {
            string id;
            if (defiedName.LocalSheet == null)
            {
                id = defiedName.Name;
            }
            else
            {
                id = defiedName.Name + "@" + ParserUtils.ToString(defiedName.LocalSheet.GetHashCode());
            }
            return id;
        }


        private static Dictionary<string, HashSet<string>> MapRidsToDefinedNames(List<string> rids, IReadOnlyList<DefinedName> definedNames)
        {
            Dictionary<string, HashSet<string>> map = new Dictionary<string, HashSet<string>>();
            if (definedNames.Count == 0)
            {
                return map;
            }
            List<string> ids = new List<string>();
            for (int i = 0; i < rids.Count; i++)
            {
                ids.Add("[" + ParserUtils.ToString(i) + "]");
            }
            foreach (DefinedName definedName in definedNames)
            {
                if (definedName.HasExternalReferences)
                {
                    foreach (string id in ids)
                    {
                        if (ExternalLinkFormulaUtils.DetectExternalLinkId(definedName.TextValue, id))
                        {
                            if (!map.TryGetValue(definedName.Name, out HashSet<string> value))
                            {
                                value = new HashSet<string>();
                                map.Add(definedName.Name, value);
                            }
                            value.Add(id);
                        }
                    }
                }
            }
            return map;
        }

    }
}

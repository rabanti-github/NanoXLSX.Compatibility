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
            List<ExternalDefinedNameReference> rawDefinedNames = Workbook.AuxiliaryData.GetData<List<ExternalDefinedNameReference>>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY);
            if (externalReferenceRids == null)
            {
                return; // No (valid) external links in workbook
            }
            Dictionary<string, ExternalLink> externalReferences = new Dictionary<string, ExternalLink>();
            Dictionary<string, DefinedName> replacementMap = new Dictionary<string, DefinedName>();
            Dictionary<string, string> structuralReplacements = new Dictionary<string, string>();
            PrepareUpdateDefinedNames(
                externalLinks,
                externalReferenceRids,
                externalReferences,
                replacementMap,
                structuralReplacements,
                rawDefinedNames);
            RebuildDefinedNames(structuralReplacements, replacementMap);
            if (Workbook.Features.ContainsWorksheetFormulas && externalReferences.Count != 0 && replacementMap.Count > 0)
            {
                UpdateCellFormulas(externalReferences, replacementMap);
            }

        }

        private void PrepareUpdateDefinedNames(
            List<ExternalLink> externalLinks,
            List<string> externalReferenceRids,
            Dictionary<string, ExternalLink> externalReferences,
            Dictionary<string, DefinedName> replacementMap,
            Dictionary<string, string> structuralReplacements,
            List<ExternalDefinedNameReference> rawDefinedNames)
        {
            for (int i = 0; i < externalReferenceRids.Count; i++)
            {
                ExternalLink link = externalLinks.FirstOrDefault(l => l != null && externalReferenceRids[i].Equals(l.WorkbookRId, StringComparison.OrdinalIgnoreCase));
                if (link == null)
                {
                    throw new IOException("Mismatch of read external links and references in workbooks detected. External references cannot be resolved.");
                }
                string rId = GetExternalLinkId(i);
                externalReferences[rId] = link;
            }
            IReadOnlyList<DefinedName> definedNames = Workbook.GetDefinedNames();
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName defiedName = definedNames[i];
                string expression = GetDefinedNameExpression(defiedName, rawDefinedNames);
                if (ExternalLinkFormulaUtils.DetectExternalLinkId(expression))
                {
                    string replacedExpression = ExternalLinkFormulaUtils.ReplaceExternalLinkId(expression, externalReferences);
                    if (ExternalLinkFormulaUtils.DetectExternalLinkId(replacedExpression))
                    {
                        throw new IOException("Mismatch of read external links and references in defined names detected. External references cannot be resolved.");
                    }
                    string id = GetDefinedNameId(defiedName);
                    if (defiedName.Type == DefinedName.NameType.Formula)
                    {
                        defiedName.ReplaceExpression(replacedExpression);
                        replacementMap[id] = defiedName;
                    }
                    else
                    {
                        structuralReplacements[id] = replacedExpression;
                    }
                }
            }
        }

        private string GetDefinedNameExpression(DefinedName definedName, List<ExternalDefinedNameReference> rawDefinedNames)
        {
            if (rawDefinedNames != null)
            {
                foreach (ExternalDefinedNameReference rawDefinedName in rawDefinedNames)
                {
                    if (string.Equals(rawDefinedName.Name, definedName.Name, StringComparison.OrdinalIgnoreCase) &&
                        IsSameScope(rawDefinedName.LocalSheetIndex, definedName.LocalSheet))
                    {
                        return rawDefinedName.Expression;
                    }
                }
            }
            return definedName.TextValue;
        }

        private bool IsSameScope(int? localSheetIndex, Worksheet localSheet)
        {
            if (!localSheetIndex.HasValue)
            {
                return localSheet == null;
            }
            int index = localSheetIndex.Value;
            return index >= 0 && index < Workbook.Worksheets.Count && object.ReferenceEquals(Workbook.Worksheets[index], localSheet);
        }

        private void UpdateCellFormulas(Dictionary<string, ExternalLink> externalReferences, Dictionary<string, DefinedName> replacementMap)
        {
            foreach (Worksheet worksheet in Workbook.Worksheets)
            {
                foreach (KeyValuePair<string, Cell> cell in worksheet.Cells)
                {
                    if (cell.Value.DataType == Cell.CellType.Formula && cell.Value.Formula != null)
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

        private void RebuildDefinedNames(Dictionary<string, string> structuralReplacements, Dictionary<string, DefinedName> replacementMap)
        {
            if (structuralReplacements.Count == 0)
            {
                return;
            }

            List<DefinedName> definedNames = Workbook.GetDefinedNames().ToList();
            for (int i = definedNames.Count - 1; i >= 0; i--)
            {
                DefinedName definedName = definedNames[i];
                Workbook.RemoveDefinedName(definedName.Name, false, definedName.LocalSheet);
            }

            foreach (DefinedName definedName in definedNames)
            {
                string id = GetDefinedNameId(definedName);
                if (structuralReplacements.TryGetValue(id, out string expression))
                {
                    DefinedName replacement = new DefinedName(
                        Workbook,
                        DefinedName.NameType.Formula,
                        definedName.Name,
                        expression,
                        null,
                        definedName.LocalSheet,
                        definedName.Comment,
                        true);
                    Workbook.AddDefinedName(replacement);
                    replacementMap[id] = replacement;
                }
                else
                {
                    definedName.Features.Add(Workbook.Features);
                    Workbook.AddDefinedName(definedName);
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

        /// <summary>
        /// Gets the one-based external workbook identifier used in formula expressions.
        /// </summary>
        private static string GetExternalLinkId(int zeroBasedIndex)
        {
            return "[" + ParserUtils.ToString(zeroBasedIndex + 1) + "]";
        }

    }
}

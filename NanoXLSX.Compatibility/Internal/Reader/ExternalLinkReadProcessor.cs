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
using System.Text;

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
                    if (DetectExternalLinkId(defiedName.TextValue))
                    {
                        string replacedExpression = ReplaceExternalLinkId(defiedName.TextValue, externalReferences);
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
                            if (DetectExternalLinkId(originalValue))
                            {
                                string replacement = ReplaceExternalLinkId(originalValue, externalReferences);
                                FormulaData formulaData = new FormulaData(replacement); // No cached value available
                                formulaData.HasExternalReferences = true;
                                cell.Value.Formula = formulaData;
                                cell.Value.Value = replacement;

                            }
                        }
                        else if (cell.Value.Formula != null)
                        {
                            string originalValue = cell.Value.Value?.ToString() ?? string.Empty;
                            if (DetectExternalLinkId(originalValue))
                            {
                                string replacement = ReplaceExternalLinkId(originalValue, externalReferences);
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


        /// <summary>
        /// Determines whether an expression contains any internal external workbook identifier such as "[1]".
        /// </summary>
        /// <param name="expression"> Formula or defined-name expression to inspect. </param>
        /// <returns>
        /// True if the expression contains an external workbook identifier; otherwise false.
        /// </returns>
        internal static bool DetectExternalLinkId(string expression)
        {
            return DetectExternalLinkId(expression, null);
        }

        /// <summary>
        /// Determines whether an expression contains an internal external workbook identifier such as "[1]".
        ///
        /// Identifiers inside Excel string constants are ignored. Structured references such as "Table1[1]" are not treated as external workbook identifiers.
        /// </summary>
        /// <param name="expression">Formula or defined-name expression to inspect.</param>
        /// <param name="targetId">Optional external workbook identifier, including its brackets, for example "[2]".</param>
        /// <returns>
        /// True if the expression contains an external workbook identifier; otherwise false.
        /// </returns>
        internal static bool DetectExternalLinkId(string expression, string targetId)
        {
            if (targetId != null &&
                !ParserUtils.IsValidExternalLinkId(targetId))
            {
                throw new ArgumentException(
                    $"The target ID '{targetId}' is not a valid external link ID.",
                    nameof(targetId));
            }

            if (string.IsNullOrEmpty(expression))
            {
                return false;
            }

            int firstOpeningBracket = expression.IndexOf('[');

            if (firstOpeningBracket < 0)
            {
                return false;
            }

            // Additional fast path when a particular ID is requested.
            //
            // This check alone is not sufficient because the occurrence could be inside a string constant.
            // It only allows an early exit when the ID is not present at all.
            if (targetId != null &&
                expression.IndexOf(
                    targetId,
                    StringComparison.Ordinal) < 0)
            {
                return false;
            }

            // Start at the first relevant character.
            // If a quote occurs before the first bracket, it must be processed so that brackets inside a string constant are ignored.
            int firstQuote = expression.IndexOf('"');

            int scanStart =
                firstQuote >= 0 &&
                firstQuote < firstOpeningBracket
                    ? firstQuote
                    : firstOpeningBracket;

            bool insideStringConstant = false;

            for (int i = scanStart; i < expression.Length; i++)
            {
                char current = expression[i];

                if (current == '"')
                {
                    if (insideStringConstant &&
                        i + 1 < expression.Length &&
                        expression[i + 1] == '"')
                    {
                        // Escaped quote inside an Excel string constant: "Text ""quoted"" text"
                        i++;
                        continue;
                    }

                    insideStringConstant = !insideStringConstant;
                    continue;
                }

                if (insideStringConstant ||
                    current != '[')
                {
                    continue;
                }

                if (!ParserUtils.TryReadExternalLinkId(
                        expression,
                        i,
                        out int identifierLength))
                {
                    continue;
                }

                if (targetId == null)
                {
                    return true;
                }

                if (identifierLength == targetId.Length &&
                    string.CompareOrdinal(
                        expression,
                        i,
                        targetId,
                        0,
                        identifierLength) == 0)
                {
                    return true;
                }

                // Skip the already validated identifier.
                i += identifierLength - 1;
            }

            return false;
        }

        /// <summary>
        /// Replaces internal external workbook identifiers with their preferred human-readable reference.
        /// Identifiers inside Excel string constants are not replaced. Unknown identifiers are left unchanged.
        /// The input parameters are assumed to be checked / non-null prior.
        /// </summary>
        /// <param name="expression">Formula or defined-name expression to process.</param>
        /// <param name="links">External links indexed by their internal identifiers, for example "[1]".</param>
        /// <returns>
        /// The transformed expression. If no replacement is necessary, the original string instance is returned.
        /// </returns>
        internal static string ReplaceExternalLinkId(string expression, Dictionary<string, ExternalLink> links)
        {
            if (expression.Length == 0 || links.Count == 0)
            {
                return expression;
            }

            int firstOpeningBracket = expression.IndexOf('[');
            if (firstOpeningBracket < 0)
            {
                return expression;
            }

            bool insideStringConstant = false;
            int unchangedSectionStart = 0;
            StringBuilder builder = null;

            for (int i = firstOpeningBracket; i < expression.Length; i++)
            {
                char current = expression[i];

                if (current == '"')
                {
                    if (insideStringConstant &&
                        i + 1 < expression.Length &&
                        expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }

                    insideStringConstant = !insideStringConstant;
                    continue;
                }

                if (insideStringConstant || current != '[')
                {
                    continue;
                }

                if (!ParserUtils.TryReadExternalLinkId(
                    expression,
                    i,
                    out int identifierLength))
                {
                    continue;
                }

                string identifier =
                    expression.Substring(i, identifierLength);

                if (!links.TryGetValue(
                    identifier,
                    out ExternalLink externalLink))
                {
                    // Non-destructive behavior: unresolved IDs remain in the expression.
                    continue;
                }

                if (externalLink == null)
                {
                    throw new ArgumentException(
                        $"The external link '{identifier}' is null.",
                        nameof(links));
                }

                string replacement = externalLink.ReadableReferenceToken;

                if (string.IsNullOrEmpty(replacement))
                {
                    throw new ArgumentException(
                        $"The external link '{identifier}' does not provide a usable URI.",
                        nameof(links));
                }

                if (builder == null)
                {
                    builder = new StringBuilder(expression.Length + replacement.Length);
                }

                builder.Append(
                    expression,
                    unchangedSectionStart,
                    i - unchangedSectionStart);

                builder.Append(replacement);

                i += identifierLength - 1;
                unchangedSectionStart = i + 1;
            }

            if (builder == null)
            {
                return expression;
            }

            builder.Append(
                expression,
                unchangedSectionStart,
                expression.Length - unchangedSectionStart);

            return builder.ToString();
        }

        private Dictionary<string, HashSet<string>> MapRidsToDefinedNames(List<string> rids, IReadOnlyList<DefinedName> definedNames)
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
                        if (DetectExternalLinkId(definedName.TextValue, id))
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

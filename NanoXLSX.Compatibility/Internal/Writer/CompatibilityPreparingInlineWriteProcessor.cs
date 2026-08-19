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
using System;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Internal.Writer
{
    /// <summary>
    /// Class responsible to prepare the workbook for its compatibility features to be written to a XLSX file
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "MAIN_COMPATIBILITY_WRITE_INLINE_PREPARATION_PROCESSOR", QueueUUID = PlugInUUID.PreparingInlineProcessor, PlugInOrder = 2000)]
    internal class CompatibilityPreparingInlineWriteProcessor : IPluginInlineWriteProcessor
    {
        #region privateFields
        private Dictionary<int, Dictionary<string, ExternalLinkResolution>> resolvedFormulas;
        private Dictionary<int, ExternalLinkResolution> resolvedDefinedNames;
        #endregion
        #region properties
        /// <summary>
        /// Write context
        /// </summary>
        public IWriteContext WriteContext { get; set; }
        #endregion
        #region methods
        /// <summary>
        /// Initializing method (interface implementation)
        /// </summary>
        /// <param name="context">Write context</param>
        public void Init(IWriteContext context)
        {
            this.WriteContext = context;
        }

        /// <summary>
        /// Main execution method of the preparing processor (interface implementation)
        /// </summary>
        public void Execute()
        {
            resolvedFormulas = new Dictionary<int, Dictionary<string, ExternalLinkResolution>>();
            resolvedDefinedNames = new Dictionary<int, ExternalLinkResolution>();

            if (!WriteContext.Workbook.Features.ContainsExternalLinks)
            {
                return; // No external links to process
            }
            List<ExternalLink> externalLinks = WriteContext.Workbook.AuxiliaryData
                .GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY)
                .OfType<ExternalLink>()
                .ToList(); // Returns a null-free list
            List<ExternalLinkCandidate> candidates = CreateExternalLinkCandidates(externalLinks);
            if (WriteContext.Workbook.Features.ContainsWorksheetFormulas)
            {
                ResolveExternalLinksFromFormulas(candidates);
            }
            if (WriteContext.Workbook.Features.ContainsDefinedNameFormulas)
            {
                ResolveExternalLinksFromDefinedNames(candidates);
            }
            StoreResolutions();
            // TODO If other resources contains possibly external links, add further handling here
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in formula cells back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="candidates">External-link formula candidates</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with 
        /// <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID, 
        /// <see cref="CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY"/> as entity ID,
        /// in a dictionary keyed first by worksheet index and then by cell address.</remarks>
        private void ResolveExternalLinksFromFormulas(List<ExternalLinkCandidate> candidates)
        {
            for (int worksheetIndex = 0; worksheetIndex < WriteContext.Workbook.Worksheets.Count; worksheetIndex++)
            {
                Worksheet worksheet = WriteContext.Workbook.Worksheets[worksheetIndex];
                if (!worksheet.Features.ContainsWorksheetFormulas || !worksheet.Features.ContainsExternalLinks)
                {
                    continue; // No external links on this worksheet
                }
                foreach (Cell cell in worksheet.CellValues)
                {
                    if (cell.DataType != Cell.CellType.Formula || cell.Formula == null || !cell.Formula.Features.ContainsExternalLinks)
                    {
                        continue; // No formula or no external link in formula
                    }

                    string expression = GetFormulaExpression(cell);
                    string cellAddress = Cell.ResolveCellAddress(cell.ColumnNumber, cell.RowNumber);
                    ExternalLinkResolution result = ResolveExpression(
                        expression,
                        new SourceInfo("cell formula", worksheet.SheetName + "!" + cellAddress),
                        candidates);
                    if (result == null)
                    {
                        continue;
                    }
                    // Note: The worksheet writer or will pass the 1-based sheetID, not the 0-based index
                    if (!resolvedFormulas.TryGetValue(worksheet.SheetID, out Dictionary<string, ExternalLinkResolution> worksheetResolutions))
                    {
                        worksheetResolutions = new Dictionary<string, ExternalLinkResolution>();
                        resolvedFormulas.Add(worksheet.SheetID, worksheetResolutions);
                    }
                    worksheetResolutions[cellAddress] = result;
                }
            }
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in defined names back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="candidates">External-link formula candidates</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with 
        /// <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID, 
        /// <see cref="CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY"/> as entity ID,
        /// in a dictionary keyed by the defined-name index.</remarks>
        private void ResolveExternalLinksFromDefinedNames(List<ExternalLinkCandidate> candidates)
        {
            IReadOnlyList<DefinedName> definedNames = WriteContext.Workbook.GetDefinedNames();
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName definedName = definedNames[i];
                if (!definedName.Features.ContainsExternalLinks)
                {
                    continue;
                }
                ExternalLinkResolution result = ResolveExpression(
                    definedName.TextValue,
                    new SourceInfo("defined name", definedName.Name),
                    candidates);
                if (result != null)
                {
                    resolvedDefinedNames[i] = result;
                }
            }
        }

        /// <summary>
        /// Gets the formula expression of a formula cell.
        /// </summary>
        /// <param name="cell">Cell to check</param>
        /// <returns>Expression of the formula</returns>
        private static string GetFormulaExpression(Cell cell)
        {
            if (cell.Formula.DefinedNameReference != null)
            {
                return cell.Formula.DefinedNameReference.Name;
            }
            return cell.Formula.Expression;
        }

        /// <summary>
        /// Stores all resolved external links in auxiliary data for later write processing
        /// </summary>
        private void StoreResolutions()
        {
            WriteContext.Workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY,
                resolvedFormulas);
            WriteContext.Workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY,
                resolvedDefinedNames);
        }

        /// <summary>
        /// Method to resolve eternal links from a Excel expression (defined name or cell formula), using the prepared candidates
        /// </summary>
        /// <param name="expression">Raw defined name or cell formula expression</param>
        /// <param name="sourceInfo">Info object with human readable texts for exception outputs</param>
        /// <param name="candidates"></param>
        /// <returns>Resolved external link object, or null if no external link is in the expression</returns>
        private static ExternalLinkResolution ResolveExpression(
            string expression,
            SourceInfo sourceInfo,
            List<ExternalLinkCandidate> candidates)
        {
            if (ExternalLinkFormulaUtils.DetectExternalLinkId(expression))
            {
                throw GetUnsupportedNumericReference(sourceInfo);
            }

            ValidateCaseInsensitiveAmbiguities(expression, sourceInfo, candidates);

            string resolvedExpression = expression;
            HashSet<int> matchedIndexes = new HashSet<int>();
            foreach (ExternalLinkCandidate candidate in candidates)
            {
                if (!TryReplaceOutsideStringConstants(
                    resolvedExpression,
                    candidate.Text,
                    "[" + ParserUtils.ToString(candidate.LinkIndexes[0]) + "]",
                    out string replacedExpression))
                {
                    continue;
                }
                if (candidate.LinkIndexes.Count > 1)
                {
                    throw GetAmbiguousReference(sourceInfo, candidate.Text, candidate.LinkIndexes);
                }

                int linkIndex = candidate.LinkIndexes[0];
                resolvedExpression = replacedExpression;
                matchedIndexes.Add(linkIndex);
            }

            if (ExternalLinkFormulaUtils.TryFindUnresolvedExternalLink(resolvedExpression, out string unresolvedToken))
            {
                throw GetUnregisteredReference(sourceInfo, unresolvedToken);
            }

            if (matchedIndexes.Count == 0)
            {
                return null;
            }
            return new ExternalLinkResolution(resolvedExpression, matchedIndexes.OrderBy(index => index).ToList());
        }

        /// <summary>
        /// Validates a source expression for clearly identifiable external links
        /// </summary>
        /// <param name="expression">Source expression</param>
        /// <param name="sourceInfo">Info object with human readable texts for exception outputs</param>
        /// <param name="candidates">List of possible external links</param>
        private static void ValidateCaseInsensitiveAmbiguities(
            string expression,
            SourceInfo sourceInfo,
            List<ExternalLinkCandidate> candidates)
        {
            foreach (IGrouping<string, ExternalLinkCandidate> group in candidates.GroupBy(
                candidate => candidate.Text,
                StringComparer.OrdinalIgnoreCase))
            {
                List<int> indexes = group
                    .SelectMany(candidate => candidate.LinkIndexes)
                    .Distinct()
                    .OrderBy(index => index)
                    .ToList();
                if (indexes.Count < 2)
                {
                    continue;
                }

                string representative = group.First().Text;
                int position = 0;
                while ((position = IndexOfOutsideStringConstants(
                    expression,
                    representative,
                    position,
                    StringComparison.OrdinalIgnoreCase)) >= 0)
                {
                    string matchedText = expression.Substring(position, representative.Length);
                    bool hasExactCandidate = group.Any(candidate =>
                        string.Equals(candidate.Text, matchedText, StringComparison.Ordinal));
                    if (!hasExactCandidate)
                    {
                        throw GetAmbiguousReference(sourceInfo, matchedText, indexes);
                    }
                    position += representative.Length;
                }
            }
        }

        /// <summary>
        /// Method to create an exception if the external links could not be clearly identified from a source (defined name or cell formula)
        /// </summary>
        /// <param name="sourceInfo">Info object with human readable texts for exception outputs</param>
        /// <param name="matchedText">Text that contains ambiguities</param>
        /// <param name="linkIndexes">Indices with indices of link candidates</param>
        /// <returns>Returns a <see cref="NotSupportedContentException"/></returns>
        private static NotSupportedContentException GetAmbiguousReference(SourceInfo sourceInfo, string matchedText, IEnumerable<int> linkIndexes)
        {
            string indexes = string.Join(", ", linkIndexes.Select(ParserUtils.ToString));
            return new NotSupportedContentException(
                "The " + sourceInfo.SourceKind + " '" + sourceInfo.SourceIdentifier + "' contains the ambiguous external link '" +
                matchedText + "', which matches external link indexes " + indexes + ".");
        }

        /// <summary>
        /// Creates an exception for a user-entered numeric external-link identifier.
        /// </summary>
        private static NotSupportedContentException GetUnsupportedNumericReference(SourceInfo sourceInfo)
        {
            return new NotSupportedContentException(
                "The " + sourceInfo.SourceKind + " '" + sourceInfo.SourceIdentifier +
                "' contains a numeric OOXML external-link identifier. Use a human-readable external workbook reference and register the workbook with WorkbookExtensions.AddExternalLink before saving.");
        }

        /// <summary>
        /// Creates an exception for a human-readable external link which is not registered on the workbook.
        /// </summary>
        private static NotSupportedContentException GetUnregisteredReference(SourceInfo sourceInfo, string unresolvedToken)
        {
            return new NotSupportedContentException(
                "The " + sourceInfo.SourceKind + " '" + sourceInfo.SourceIdentifier + "' contains the unregistered external link '" +
                unresolvedToken + "'. Register the external workbook with WorkbookExtensions.AddExternalLink before saving.");
        }

        /// <summary>
        /// Replaces all ordinal occurrences outside Excel string constants.
        /// </summary>
        private static bool TryReplaceOutsideStringConstants(
            string expression,
            string oldValue,
            string newValue,
            out string result)
        {
            System.Text.StringBuilder builder = null;
            bool insideStringConstant = false;
            int unchangedSectionStart = 0;
            for (int i = 0; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current == '"')
                {
                    if (insideStringConstant && i + 1 < expression.Length && expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    insideStringConstant = !insideStringConstant;
                    continue;
                }
                if (insideStringConstant || i + oldValue.Length > expression.Length ||
                    string.CompareOrdinal(expression, i, oldValue, 0, oldValue.Length) != 0 ||
                    !HasValidCandidatePrefix(expression, i, oldValue))
                {
                    continue;
                }

                if (builder == null)
                {
                    builder = new System.Text.StringBuilder(expression.Length);
                }
                builder.Append(expression, unchangedSectionStart, i - unchangedSectionStart);
                builder.Append(newValue);
                i += oldValue.Length - 1;
                unchangedSectionStart = i + 1;
            }

            if (builder == null)
            {
                result = expression;
                return false;
            }
            builder.Append(expression, unchangedSectionStart, expression.Length - unchangedSectionStart);
            result = builder.ToString();
            return true;
        }

        /// <summary>
        /// Finds an occurrence outside Excel string constants.
        /// </summary>
        private static int IndexOfOutsideStringConstants(
            string expression,
            string value,
            int startIndex,
            StringComparison comparison)
        {
            bool insideStringConstant = false;
            for (int i = 0; i + value.Length <= expression.Length; i++)
            {
                char current = expression[i];
                if (current == '"')
                {
                    if (insideStringConstant && i + 1 < expression.Length && expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    insideStringConstant = !insideStringConstant;
                    continue;
                }
                if (i >= startIndex && !insideStringConstant &&
                    string.Compare(expression, i, value, 0, value.Length, comparison) == 0 &&
                    HasValidCandidatePrefix(expression, i, value))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// Prevents a filename-only candidate from matching the tail of a different explicit path.
        /// </summary>
        private static bool HasValidCandidatePrefix(string expression, int matchIndex, string candidate)
        {
            if (matchIndex == 0 || candidate.Length == 0 || candidate[0] != '[')
            {
                return true;
            }

            char previous = expression[matchIndex - 1];
            return !char.IsLetterOrDigit(previous) && previous != '_' && previous != '.' &&
                previous != ':' && previous != '/' && previous != '\\';
        }

        /// <summary>
        /// Method to identify possible candidates of external links from a unresolved expression
        /// </summary>
        /// <param name="externalLinks">List of external link objects</param>
        /// <returns>List of not yet validated candidates</returns>
        private static List<ExternalLinkCandidate> CreateExternalLinkCandidates(List<ExternalLink> externalLinks)
        {
            Dictionary<string, HashSet<int>> candidates = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);

            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink externalLink = externalLinks[i];
                int linkIndex = i + 1;
                HashSet<string> linkCandidates = new HashSet<string>(StringComparer.Ordinal);
                foreach (string uri in externalLink.GetWorkbookLocations())
                {
                    foreach (string candidate in CreateUriCandidates(uri))
                    {
                        linkCandidates.Add(candidate);
                    }
                }
                string readableToken = externalLink.ReadableReferenceToken;
                if (!string.IsNullOrEmpty(readableToken))
                {
                    linkCandidates.Add(readableToken);
                    linkCandidates.Add(readableToken.Replace('\\', '/'));
                    linkCandidates.Add(readableToken.Replace('/', '\\'));
                }
                foreach (string candidate in linkCandidates)
                {
                    if (!candidates.TryGetValue(candidate, out HashSet<int> indexes))
                    {
                        indexes = new HashSet<int>();
                        candidates[candidate] = indexes;
                    }
                    indexes.Add(linkIndex);
                }
            }

            return candidates
                .Select(pair => new ExternalLinkCandidate(
                    pair.Key,
                    pair.Value.OrderBy(index => index).ToList()))
                .OrderByDescending(candidate => candidate.Text.Length)
                .ThenBy(candidate => candidate.Text, StringComparer.Ordinal)
                .ToList();
        }

        /// <summary>
        /// Analyzes a URI or path and creates possible representations used in Excel formulas and defined names.
        /// </summary>
        /// <param name="uriText">Raw URI or path text.</param>
        /// <returns>A set containing possible formula path representations.</returns>
        private static HashSet<string> CreateUriCandidates(string uriText)
        {
            HashSet<string> candidates = new HashSet<string>(StringComparer.Ordinal);
            string value = uriText.Trim(); // Should already be sanitized

            bool isAbsoluteUri = Uri.TryCreate(value, UriKind.Absolute, out Uri uri);

            if (isAbsoluteUri && uri.IsFile)
            {
                string filePath = GetFilePathCandidate(uri);
                AddPathCandidates(candidates, filePath, true);
                return candidates;
            }

            // For relative paths or strings that are not valid absolute URIs, create slash and backslash aliases.
            // For absolute non-file URIs, preserve the original separator form.
            bool addSeparatorAliases = !isAbsoluteUri;
            AddPathCandidates(candidates, value, addSeparatorAliases);
            return candidates;
        }

        /// <summary>
        /// Converts a file URI to one normalized path representation.
        /// </summary>
        /// <param name="uri">Absolute file URI.</param>
        /// <returns>
        /// The normalized path, or null if the URI does not contain a file path.
        /// </returns>
        private static string GetFilePathCandidate(Uri uri)
        {
            // AbsolutePath is used instead of processing both LocalPath and AbsolutePath.
            // This avoids generating mostly redundant candidates. AbsolutePath is still URI-escaped and must therefore be decoded.
            string path = Uri.UnescapeDataString(uri.AbsolutePath).Replace('\\', '/');

            bool isRemoteFile = !string.IsNullOrEmpty(uri.Host) && !uri.IsLoopback;

            if (isRemoteFile)
            {
                string remotePath = path.TrimStart('/');

                return "//" + uri.Host + "/" + remotePath;
            }

            // A Windows drive path in a file URI commonly has this form: "/C:/directory/file.xlsx". Remove the URI-specific leading slash.
            if (path.Length >= 3 && path[0] == '/' && char.IsLetter(path[1]) && path[2] == ':')
            {
                path = path.Substring(1);
            }
            return path;
        }

        /// <summary>
        /// Adds a possible formula path to the set of candidates.
        /// </summary>
        /// <param name="candidates">Target candidate collection.</param>
        /// <param name="path">Raw path expression.</param>
        /// <param name="addSeparatorAliases">
        /// If true, variants using forward and backward slashes are added.
        /// </param>
        private static void AddPathCandidates(HashSet<string> candidates, string path, bool addSeparatorAliases)
        {
            string formulaPath = ToFormulaPath(path);

            candidates.Add(formulaPath);
            if (!addSeparatorAliases)
            {
                return;
            }
            candidates.Add(formulaPath.Replace('\\', '/'));
            candidates.Add(formulaPath.Replace('/', '\\'));
        }

        /// <summary>
        /// Converts a path to its Excel formula representation.
        /// </summary>
        /// <param name="path">Path to convert.</param>
        /// <returns>
        /// The converted path, or null if the value does not contain a filename.
        /// </returns>
        private static string ToFormulaPath(string path)
        {
            string value = path.Trim();

            int separator = Math.Max(value.LastIndexOf('/'), value.LastIndexOf('\\'));

            string directory;
            string filename;

            if (separator >= 0)
            {
                directory = value.Substring(0, separator + 1);
                filename = value.Substring(separator + 1);
            }
            else
            {
                directory = string.Empty;
                filename = value;
            }

            if (filename[0] == '[' && filename[filename.Length - 1] == ']')
            {
                return value;
            }

            return directory + "[" + filename + "]";
        }
        #endregion

        #region helperClasses

        /// <summary>
        /// Helper class, representing an expression possibly containing one or many external links as full text
        /// </summary>
        private sealed class ExternalLinkCandidate
        {
            /// <summary>
            /// Unresolved text / expression
            /// </summary>
            public string Text { get; }
            // Indices where external links may start and end
            public List<int> LinkIndexes { get; }

            /// <summary>
            /// Constructor with parameters
            /// </summary>
            /// <param name="text">Unresolved text</param>
            /// <param name="linkIndexes">Identified indices</param>
            public ExternalLinkCandidate(string text, List<int> linkIndexes)
            {
                Text = text;
                LinkIndexes = linkIndexes;
            }
        }

        /// <summary>
        /// Helper class, holding verbose source info for clearer exception messages
        /// </summary>
        private sealed class SourceInfo
        {
            /// <summary>
            /// Human readable identifier of the expressions origin
            /// </summary>
            public string SourceKind { get; }
            /// <summary>
            /// Defined name ID or cell origin
            /// </summary>
            public string SourceIdentifier { get; }

            /// <summary>
            /// Constructor with parameters
            /// </summary>
            /// <param name="sourceKind">Human readable identifier of the expressions origin</param>
            /// <param name="sourceIdentifier">Defined name ID or cell origin</param>
            public SourceInfo(string sourceKind, string sourceIdentifier)
            {
                this.SourceKind = sourceKind;
                this.SourceIdentifier = sourceIdentifier;
            }
        }
        #endregion

    }
}

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

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class responsible to prepare the workbook for its compatibility features to be written to a XLSX file
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "MAIN_COMPATIBILITY_WRITE_INLINE_PREPARATION_PROCESSOR", QueueUUID = PlugInUUID.PreparingInlineProcessor, PlugInOrder = 2000)]
    internal class CompatibilityPreparingInlineWriteProcessor : IPluginInlineWriteProcessor
    {
        /// <summary>
        /// Write context
        /// </summary>
        public IWriteContext WriteContext { get; set; }

        /// <summary>
        /// Initializing method
        /// </summary>
        /// <param name="context">Writ context</param>
        public void Init(IWriteContext context)
        {
            this.WriteContext = context;
        }

        /// <summary>
        /// Main execution method of the preparing processor
        /// </summary>
        public void Execute()
        {
            List<ExternalLink> externalLinks = WriteContext.Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY)
                .OfType<ExternalLink>()
                .ToList(); // Returns a null-free list
            if (externalLinks == null || externalLinks.Count == 0)
            {
                return; // No external links to process
            }
            ResolveExternalLinksFromFormulas(externalLinks);
            ResolveExternalLinksFromDefinedNames(externalLinks);
            // TODO If other resources contains possibly external links, add further handling here
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in formula cells back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="externalLinks">List of ExternalLink objects</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with 
        /// <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID, 
        /// <see cref="CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY"/> as entity ID 
        /// and the worksheet index (as string) and cell address, separated by a colon, as object ID. (e.g. "0:C3)"</remarks>
        private void ResolveExternalLinksFromFormulas(List<ExternalLink> externalLinks)
        {
            List<ExternalLinkCandidate> candidates = CreateExternalLinkCandidates(externalLinks);
            for (int worksheetIndex = 0; worksheetIndex < WriteContext.Workbook.Worksheets.Count; worksheetIndex++)
            {
                Worksheet worksheet = WriteContext.Workbook.Worksheets[worksheetIndex];
                foreach (Cell cell in worksheet.CellValues)
                {
                    if (cell.DataType != Cell.CellType.Formula)
                    {
                        continue;
                    }

                    string expression = GetFormulaExpression(cell);
                    string cellAddress = Cell.ResolveCellAddress(cell.ColumnNumber, cell.RowNumber);
                    ExternalLinkResolution result = ResolveExpression(
                        expression,
                        new SourceInfo("cell formula", worksheet.SheetName + "!" + cellAddress),
                        candidates);
                    StoreResolution(
                        CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, ParserUtils.ToString(worksheetIndex) + ":" + cellAddress, result);
                }
            }
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in defined names back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="externalLinks">List of ExternalLink objects</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with 
        /// <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID, 
        /// <see cref="CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY"/> as entity ID 
        /// and the indexer (int as string) as object ID.</remarks>
        private void ResolveExternalLinksFromDefinedNames(List<ExternalLink> externalLinks)
        {
            List<ExternalLinkCandidate> candidates = CreateExternalLinkCandidates(externalLinks);
            IReadOnlyList<DefinedName> definedNames = WriteContext.Workbook.GetDefinedNames();
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName definedName = definedNames[i];
                ExternalLinkResolution result = ResolveExpression(
                    definedName.TextValue,
                    new SourceInfo("defined name", definedName.Name),
                    candidates);
                StoreResolution(CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, ParserUtils.ToString(i), result);
            }
        }

        /// <summary>
        /// Gets the formula expression of a cell. If a <see cref="FormulaData"/> objects is not existing, the cell value will be used
        /// </summary>
        /// <param name="cell">Cell to check</param>
        /// <returns>Expression of the formula</returns>
        private static string GetFormulaExpression(Cell cell)
        {
            if (cell.Formula == null)
            {
                return cell.Value as string ?? cell.Value?.ToString();
            }
            if (cell.Formula.DefinedNameReference != null)
            {
                return cell.Formula.DefinedNameReference.Name;
            }
            return cell.Formula.Expression;
        }

        /// <summary>
        /// Stores resolved external links in auxiliary data for later write processing
        /// </summary>
        /// <param name="entityId">Grouping entity ID for resolved external links</param>
        /// <param name="valueId">ID of the actual external link object (index as string)</param>
        /// <param name="result">External link object</param>
        private void StoreResolution(string entityId, string valueId, ExternalLinkResolution result)
        {
            if (result == null)
            {
                return;
            }
            WriteContext.Workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                entityId,
                valueId,
                result
                );
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
            if (string.IsNullOrEmpty(expression) || candidates.Count == 0)
            {
                return null;
            }

            ValidateCaseInsensitiveAmbiguities(expression, sourceInfo, candidates);

            string resolvedExpression = expression;
            HashSet<int> matchedIndexes = new HashSet<int>();
            foreach (ExternalLinkCandidate candidate in candidates)
            {
                if (resolvedExpression.IndexOf(candidate.Text, StringComparison.Ordinal) < 0)
                {
                    continue;
                }
                if (candidate.LinkIndexes.Count > 1)
                {
                    throw GetAmbiguousReference(sourceInfo, candidate.Text, candidate.LinkIndexes);
                }

                int linkIndex = candidate.LinkIndexes[0];
                resolvedExpression = resolvedExpression.Replace(candidate.Text, "[" + ParserUtils.ToString(linkIndex) + "]");
                matchedIndexes.Add(linkIndex);
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
                while ((position = expression.IndexOf(representative, position, StringComparison.OrdinalIgnoreCase)) >= 0)
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
                foreach (string uri in externalLink.Uris)
                {
                    foreach (string candidate in CreateUriCandidates(uri))
                    {
                        if (!candidates.TryGetValue(candidate, out HashSet<int> indexes))
                        {
                            indexes = new HashSet<int>();
                            candidates[candidate] = indexes;
                        }
                        indexes.Add(linkIndex);
                    }
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

                // "file://server" does not identify a file. Returning null is intentional; AddPathCandidates handles it.
                if (remotePath.Length == 0)
                {
                    return null;
                }
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

            // HashSet<string> permits null values in these target frameworks. Therefore the null check must happen before candidates.Add().

            if (formulaPath == null)
            {
                return;
            }
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
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

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

            //A path ending with a separator denotes a directory or host, not a file.

            if (filename.Length == 0)
            {
                return null;
            }

            if (filename[0] == '[' && filename[filename.Length - 1] == ']')
            {
                return value;
            }

            return directory + "[" + filename + "]";
        }

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

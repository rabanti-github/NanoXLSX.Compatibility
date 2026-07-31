/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using System;
using System.Collections.Generic;
using System.Globalization;
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
            List<ExternalLink> externalLinks = WriteContext.Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, "a");
            ResolveExternalLinksFormCells(externalLinks);
            ResolveExternalLinksFormDefinedNames(externalLinks);
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in defined names back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="externalLinks">List of ExternalLink objects</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID and the indexer (int as string) as entity ID.</remarks>
        private void ResolveExternalLinksFormDefinedNames(List<ExternalLink> externalLinks)
        {
            List<ExternalLinkCandidate> candidates = CreateExternalLinkCandidates(externalLinks);
            IReadOnlyList<DefinedName> definedNames = WriteContext.Workbook.GetDefinedNames();
            for (int i = 0; i < definedNames.Count; i++)
            {
                DefinedName definedName = definedNames[i];
                ResolutionResult result = ResolveExpression(
                    definedName.TextValue,
                    "defined name",
                    definedName.Name,
                    candidates);
                StoreResolution("definedName:" + ToInvariantString(i), result);
            }
        }

        /// <summary>
        /// Method to translate external link expressions (with file name and optional path) in formula cells back to the internal indexer representation (e.g. [1])
        /// </summary>
        /// <param name="externalLinks">List of ExternalLink objects</param>
        /// \remark <remarks>The method does not overwrite the expression of the defined name. 
        /// It stores the resolved expression in <see cref="Workbook.AuxiliaryData"/> with <see cref="PlugInUUID.CompatibilityInlineProcessor"/> as plugin ID and the indexer (int as string) as entity ID.</remarks>
        private void ResolveExternalLinksFormCells(List<ExternalLink> externalLinks)
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

                    string expression = GetSerializedFormulaExpression(cell);
                    string cellAddress = Cell.ResolveCellAddress(cell.ColumnNumber, cell.RowNumber);
                    ResolutionResult result = ResolveExpression(
                        expression,
                        "cell formula",
                        worksheet.SheetName + "!" + cellAddress,
                        candidates);
                    StoreResolution(
                        "cell:" + ToInvariantString(worksheetIndex) + ":" + cellAddress,
                        result);
                }
            }
        }

        private static string GetSerializedFormulaExpression(Cell cell)
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

        private void StoreResolution(string valueId, ResolutionResult result)
        {
            if (result == null)
            {
                return;
            }
            foreach (int linkIndex in result.LinkIndexes)
            {
                WriteContext.Workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    ToInvariantString(linkIndex),
                    valueId,
                    result.Expression);
            }
        }

        private static ResolutionResult ResolveExpression(
            string expression,
            string sourceKind,
            string sourceIdentifier,
            List<ExternalLinkCandidate> candidates)
        {
            if (string.IsNullOrEmpty(expression) || candidates.Count == 0)
            {
                return null;
            }

            ValidateCaseInsensitiveAmbiguities(expression, sourceKind, sourceIdentifier, candidates);

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
                    ThrowAmbiguousReference(sourceKind, sourceIdentifier, candidate.Text, candidate.LinkIndexes);
                }

                int linkIndex = candidate.LinkIndexes[0];
                resolvedExpression = resolvedExpression.Replace(candidate.Text, "[" + ToInvariantString(linkIndex) + "]");
                matchedIndexes.Add(linkIndex);
            }

            if (matchedIndexes.Count == 0)
            {
                return null;
            }
            return new ResolutionResult(resolvedExpression, matchedIndexes.OrderBy(index => index).ToList());
        }

        private static void ValidateCaseInsensitiveAmbiguities(
            string expression,
            string sourceKind,
            string sourceIdentifier,
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
                        ThrowAmbiguousReference(sourceKind, sourceIdentifier, matchedText, indexes);
                    }
                    position += representative.Length;
                }
            }
        }

        private static void ThrowAmbiguousReference(
            string sourceKind,
            string sourceIdentifier,
            string matchedText,
            IEnumerable<int> linkIndexes)
        {
            string indexes = string.Join(", ", linkIndexes.Select(ToInvariantString));
            throw new NotSupportedContentException(
                "The " + sourceKind + " '" + sourceIdentifier + "' contains the ambiguous external link '" +
                matchedText + "', which matches external link indexes " + indexes + ".");
        }

        private static List<ExternalLinkCandidate> CreateExternalLinkCandidates(List<ExternalLink> externalLinks)
        {
            Dictionary<string, HashSet<int>> candidates = new Dictionary<string, HashSet<int>>(StringComparer.Ordinal);
            if (externalLinks == null)
            {
                return new List<ExternalLinkCandidate>();
            }

            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink externalLink = externalLinks[i];
                if (externalLink == null)
                {
                    continue;
                }
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

        private static HashSet<string> CreateUriCandidates(string uriText)
        {
            HashSet<string> paths = new HashSet<string>(StringComparer.Ordinal);
            if (string.IsNullOrWhiteSpace(uriText))
            {
                return paths;
            }

            string value = uriText.Trim();
            if (Uri.TryCreate(value, UriKind.Absolute, out Uri uri) && uri.IsFile)
            {
                string localPath = Uri.UnescapeDataString(uri.LocalPath);
                if (!string.IsNullOrEmpty(uri.Host) && !uri.IsLoopback)
                {
                    string normalizedLocalPath = localPath.Replace('\\', '/');
                    string hostPrefix = "//" + uri.Host + "/";
                    if (!normalizedLocalPath.StartsWith(hostPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        normalizedLocalPath = hostPrefix + normalizedLocalPath.TrimStart('/');
                    }
                    localPath = normalizedLocalPath;
                }
                AddPathCandidates(paths, localPath, true);

                string absolutePath = Uri.UnescapeDataString(uri.AbsolutePath);
                if (!string.IsNullOrEmpty(uri.Host) && !uri.IsLoopback)
                {
                    absolutePath = "//" + uri.Host + "/" + absolutePath.TrimStart('/');
                }
                if (absolutePath.Length >= 3 && absolutePath[0] == '/' &&
                    char.IsLetter(absolutePath[1]) && absolutePath[2] == ':')
                {
                    absolutePath = absolutePath.Substring(1);
                }
                AddPathCandidates(paths, absolutePath, true);
            }
            else
            {
                bool addSeparatorAliases = !Uri.TryCreate(value, UriKind.Absolute, out Uri absoluteUri)
                    || absoluteUri.IsFile;
                AddPathCandidates(paths, value, addSeparatorAliases);
            }
            return paths;
        }

        private static void AddPathCandidates(HashSet<string> candidates, string path, bool addSeparatorAliases)
        {
            string formulaPath = ToFormulaPath(path);
            if (string.IsNullOrEmpty(formulaPath))
            {
                return;
            }
            candidates.Add(formulaPath);
            if (addSeparatorAliases)
            {
                candidates.Add(formulaPath.Replace('\\', '/'));
                candidates.Add(formulaPath.Replace('/', '\\'));
            }
        }

        private static string ToFormulaPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }
            string value = path.Trim();
            int separator = Math.Max(value.LastIndexOf('/'), value.LastIndexOf('\\'));
            string directory = separator >= 0 ? value.Substring(0, separator + 1) : string.Empty;
            string filename = separator >= 0 ? value.Substring(separator + 1) : value;
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

        private static string ToInvariantString(int value)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        private sealed class ExternalLinkCandidate
        {
            public string Text { get; }
            public List<int> LinkIndexes { get; }

            public ExternalLinkCandidate(string text, List<int> linkIndexes)
            {
                Text = text;
                LinkIndexes = linkIndexes;
            }
        }

        private sealed class ResolutionResult
        {
            public string Expression { get; }
            public List<int> LinkIndexes { get; }

            public ResolutionResult(string expression, List<int> linkIndexes)
            {
                Expression = expression;
                LinkIndexes = linkIndexes;
            }
        }


    }
}

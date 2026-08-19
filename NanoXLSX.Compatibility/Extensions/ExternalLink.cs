/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a link to an external workbook.
    /// </summary>
    public class ExternalLink
    {
        private readonly List<ExternalWorksheet> worksheets = new List<ExternalWorksheet>();
        private readonly List<ExternalDefinedName> definedNames = new List<ExternalDefinedName>();
        /// <summary>
        /// Gets the absolute path or URI of the external workbook.
        /// </summary>
        public string AbsoluteUri { get; private set; }

        /// <summary>
        /// Gets the optional relative path or URI of the external workbook.
        /// </summary>
        public string RelativeUri { get; private set; }

        /// <summary>
        /// Gets the primary OOXML relationship target.
        /// </summary>
        internal string TargetUri { get; private set; }

        /// <summary>
        /// Gets the optional absolute alternate OOXML relationship target.
        /// </summary>
        internal string AbsoluteAlternateUri { get; private set; }

        /// <summary>
        /// Gets the optional relative alternate OOXML relationship target.
        /// </summary>
        internal string RelativeAlternateUri { get; private set; }

        /// <summary>
        /// Gets the worksheets declared for the external workbook.
        /// The worksheet index corresponds to the sheetId used in OOXML.
        /// </summary>
        public IReadOnlyList<ExternalWorksheet> Worksheets => worksheets;

        /// <summary>
        /// Gets the defined names declared by the external workbook.
        /// </summary>
        public IReadOnlyList<ExternalDefinedName> DefinedNames => definedNames;

        /// <summary>
        /// Internal Relationship ID of the external link file in the workbook definition (defines the order / indexer)
        /// </summary>
        internal string WorkbookRId { get; set; }

        /// <summary>
        /// URI, used to replace internal IDs with readable file paths
        /// </summary>
        internal string ReadableReferenceToken
        {
            get
            {
                string uri = GetReadableUri();
                if (!string.IsNullOrEmpty(uri))
                {
                    return "[" + uri + "]";
                }
                return null;
            }
        }

        /// <summary>
        /// Constructor of an external workbook link.
        /// </summary>
        internal ExternalLink()
        {
        }

        /// <summary>
        /// Constructor of an external workbook link with absolute URI. A valid file definition (e.g. 'workbook.xlsx') is mandatory in the path
        /// </summary>
        /// <param name="absoluteUri">Absolute path or URI of the external workbook.</param>
        /// \remark <remarks>No relative path is inferred when only an absolute URI is supplied.</remarks>
        public ExternalLink(string absoluteUri)
        {
            SetAuthoringUris(absoluteUri, null);
        }

        /// <summary>
        /// Constructor of an external workbook link with absolute and relative URI. A valid file definition (e.g. 'workbook.xlsx') is mandatory in relative path
        /// </summary>
        /// <param name="absoluteUri">Absolute path or URI of the external workbook.</param>
        /// <param name="relativeUri">Relative path of the external workbook</param>
        /// \remark <remarks>The relative path is usually just a filename  (e.g. 'workbook.xlsx') but it can be explicitly defined with a directory part  (e.g. '../ workbooks/workbook.xlsx') </remarks>
        public ExternalLink(string absoluteUri, string relativeUri)
        {
            SetAuthoringUris(absoluteUri, relativeUri);
        }


        /// <summary>
        /// Creates a builder for this external workbook link.
        /// </summary>
        public ExternalLinkBuilder CreateBuilder()
        {
            return new ExternalLinkBuilder(this);
        }

        /// <summary>
        /// Adds a defined name from the external workbook.
        /// </summary>
        /// <param name="name">Name of the external defined name.</param>
        /// <param name="refersTo">Expression referenced by the defined name, for example ='Sheet1'!$B$1.</param>
        public void AddDefinedName(string name, string refersTo)
        {
            AddDefinedName(new ExternalDefinedName(name, refersTo));
        }

        /// <summary>
        /// Adds a defined name from the external workbook.
        /// </summary>
        /// <param name="definedName">Defined name object</param>
        public void AddDefinedName(ExternalDefinedName definedName)
        {
            if (definedName == null)
            {
                throw new ArgumentException("An external defined name cannot be null");
            }

            if (definedNames.Any(
                    item => string.Equals(
                        item.Name,
                        definedName.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"The external defined name '{definedName.Name}' already exists.");
            }
            definedNames.Add(definedName);
        }

        /// <summary>
        /// Sets the authoring URIs of an external workbook.
        /// </summary>
        /// <param name="absoluteUri">Absolute URI of the external workbook</param>
        /// <param name="relativeUri">Optional relative URI of the external workbook</param>
        private void SetAuthoringUris(string absoluteUri, string relativeUri)
        {
            ValidateWorkbookLocation(absoluteUri, true);
            if (relativeUri != null)
            {
                ValidateWorkbookLocation(relativeUri, false);
            }

            AbsoluteUri = absoluteUri;
            RelativeUri = relativeUri;
            TargetUri = relativeUri ?? absoluteUri;
            AbsoluteAlternateUri = relativeUri == null ? null : absoluteUri;
            RelativeAlternateUri = null;
        }

        /// <summary>
        /// Sets URI roles read from an OOXML external-link part.
        /// </summary>
        internal void SetReadUris(string targetUri, string absoluteAlternateUri, string relativeAlternateUri)
        {
            ValidateWorkbookLocation(targetUri, null);
            if (absoluteAlternateUri != null)
            {
                ValidateWorkbookLocation(absoluteAlternateUri, true);
            }
            if (relativeAlternateUri != null)
            {
                ValidateWorkbookLocation(relativeAlternateUri, false);
            }

            TargetUri = targetUri;
            AbsoluteAlternateUri = absoluteAlternateUri;
            RelativeAlternateUri = relativeAlternateUri;

            bool targetIsAbsolute = IsAbsoluteWorkbookLocation(targetUri);
            AbsoluteUri = targetIsAbsolute ? targetUri : absoluteAlternateUri;
            RelativeUri = targetIsAbsolute ? relativeAlternateUri : targetUri;
        }

        /// <summary>
        /// Gets all distinct workbook locations represented by this link.
        /// </summary>
        internal IReadOnlyList<string> GetWorkbookLocations()
        {
            List<string> locations = new List<string>();
            AddDistinctLocation(locations, TargetUri);
            AddDistinctLocation(locations, AbsoluteAlternateUri);
            AddDistinctLocation(locations, RelativeAlternateUri);
            return locations.AsReadOnly();
        }

        /// <summary>
        /// Gets the relationship projection used by external-link package writers.
        /// </summary>
        internal IReadOnlyList<ExternalLinkUriRelationship> GetUriRelationships()
        {
            List<ExternalLinkUriRelationship> relationships = new List<ExternalLinkUriRelationship>();
            relationships.Add(new ExternalLinkUriRelationship(
                "rId1",
                TargetUri,
                SerializeRelationshipTarget(TargetUri),
                ExternalLinkUriRole.Target));

            int relationshipIndex = 2;
            if (AbsoluteAlternateUri != null)
            {
                relationships.Add(new ExternalLinkUriRelationship(
                    "rId" + relationshipIndex,
                    AbsoluteAlternateUri,
                    SerializeRelationshipTarget(AbsoluteAlternateUri),
                    ExternalLinkUriRole.AbsoluteAlternate));
                relationshipIndex++;
            }
            if (RelativeAlternateUri != null)
            {
                relationships.Add(new ExternalLinkUriRelationship(
                    "rId" + relationshipIndex,
                    RelativeAlternateUri,
                    SerializeRelationshipTarget(RelativeAlternateUri),
                    ExternalLinkUriRole.RelativeAlternate));
            }
            return relationships.AsReadOnly();
        }

        /// <summary>
        /// Converts a workbook location to an RFC 3986 relationship target.
        /// </summary>
        internal static string SerializeRelationshipTarget(string value)
        {
            ValidateWorkbookLocation(value, null);
            string location = value.Trim();

            if (IsWindowsDrivePath(location))
            {
                string normalized = location.Replace('\\', '/');
                string drive = normalized.Substring(0, 2);
                string path = EscapePath(normalized.Substring(2));
                return "file:///" + drive + path;
            }

            if (IsUncPath(location))
            {
                string normalized = location.Replace('\\', '/').TrimStart('/');
                int separator = normalized.IndexOf('/');
                string host = normalized.Substring(0, separator);
                string path = normalized.Substring(separator);
                return "file://" + host + EscapePath(path);
            }

            if (IsUnixRootedPath(location))
            {
                return "file://" + EscapePath(location.Replace('\\', '/'));
            }

            if (Uri.TryCreate(location.Replace('\\', '/'), UriKind.Absolute, out Uri absoluteUri))
            {
                return absoluteUri.AbsoluteUri;
            }

            string relativeLocation = location.Replace('\\', '/');
            int query = relativeLocation.IndexOf('?');
            int fragment = relativeLocation.IndexOf('#');
            int suffix = query < 0 ? fragment : fragment < 0 ? query : Math.Min(query, fragment);
            string relativePath = suffix < 0 ? relativeLocation : relativeLocation.Substring(0, suffix);
            string relativeSuffix = suffix < 0 ? string.Empty : relativeLocation.Substring(suffix);
            return EscapePath(relativePath) + relativeSuffix;
        }

        /// <summary>
        /// Gets the readable URI from several constellations
        /// </summary>
        /// <returns>Readable URI or null if none could be determined</returns>
        private string GetReadableUri()
        {
            if (TargetUri != null && IsAbsoluteWorkbookLocation(TargetUri))
            {
                return TargetUri;
            }
            if (AbsoluteAlternateUri != null)
            {
                return AbsoluteAlternateUri;
            }
            if (TargetUri != null)
            {
                return TargetUri;
            }
            return RelativeAlternateUri;
        }

        /// <summary>
        /// Add the correct URI locations for different URI types of relationships
        /// </summary>
        /// <param name="locations">Reference of the URI list</param>
        /// <param name="value">Value to assign</param>
        private static void AddDistinctLocation(List<string> locations, string value)
        {
            if (value != null && !locations.Contains(value, StringComparer.Ordinal))
            {
                locations.Add(value);
            }
        }

        /// <summary>
        /// Validates the URI of a worksheet (relationship target)
        /// </summary>
        /// <param name="value">Value (URI) to validate</param>
        /// <param name="mustBeAbsolute">If true, the validated URI must be absolute</param>
        /// <exception cref="ArgumentException">Thrown if the URI is invalid as URI for a external workbook</exception>
        private static void ValidateWorkbookLocation(string value, bool? mustBeAbsolute)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new ArgumentException("The URI cannot be null or empty.");
            }

            bool isAbsolute = IsAbsoluteWorkbookLocation(value.Trim());
            if (mustBeAbsolute == true && !isAbsolute)
            {
                throw new ArgumentException($"The URI must be an absolute workbook location: '{value}'.");
            }
            if (mustBeAbsolute == false && isAbsolute)
            {
                throw new ArgumentException($"The URI must be a relative workbook location: '{value}'.");
            }

            string path = GetLocationPath(value.Trim(), isAbsolute);
            int separator = Math.Max(path.LastIndexOf('/'), path.LastIndexOf('\\'));
            string filename = separator < 0 ? path : path.Substring(separator + 1);
            int extensionSeparator = filename.LastIndexOf('.');
            if (filename.Length == 0 || extensionSeparator <= 0 || extensionSeparator == filename.Length - 1)
            {
                throw new ArgumentException($"The URI must point to a file with an extension: '{value}'.");
            }
        }

        /// <summary>
        /// Gets the escaped path of a URI, used as relationship target
        /// </summary>
        /// <param name="value">Value (raw URI) to be processed</param>
        /// <param name="isAbsolute">If true, the URI is expected as absolute path</param>
        /// <returns>Escaped location path</returns>
        private static string GetLocationPath(string value, bool isAbsolute)
        {
            if (isAbsolute && !IsWindowsDrivePath(value) && !IsUncPath(value) && !IsUnixRootedPath(value)
                && Uri.TryCreate(value.Replace('\\', '/'), UriKind.Absolute, out Uri absoluteUri))
            {
                return Uri.UnescapeDataString(absoluteUri.AbsolutePath);
            }

            int query = value.IndexOf('?');
            int fragment = value.IndexOf('#');
            int suffix = fragment < 0 ? query < 0 ? fragment : query : query < 0 ? fragment : Math.Min(query, fragment);
            string path = suffix < 0 ? value : value.Substring(0, suffix);
            return Uri.UnescapeDataString(path);
        }

        /// <summary>
        /// Method to escape an URI path
        /// </summary>
        /// <param name="path">URI to escape</param>
        /// <returns>Escaped URI</returns>
        private static string EscapePath(string path)
        {
            string[] segments = path.Split('/');
            for (int i = 0; i < segments.Length; i++)
            {
                segments[i] = Uri.EscapeDataString(Uri.UnescapeDataString(segments[i]));
            }
            return string.Join("/", segments);
        }


        /// <summary>
        /// Adds an external worksheet
        /// </summary>
        /// <param name="worksheet">External Worksheet to Add</param>
        internal void AddWorksheet(ExternalWorksheet worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentException("The worksheet cannot be null");
            }

            if (worksheets.Any(item => string.Equals(item.Name, worksheet.Name, StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"The external worksheet '{worksheet.Name}' already exists.");
            }
            worksheets.Add(worksheet);
        }

        /// <summary>
        /// Gets an external worksheet by name.
        /// </summary>
        /// <param name="name">Name of the external defined name</param>
        public ExternalWorksheet GetWorksheet(string name)
        {
            ExternalWorksheet worksheet = worksheets.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            if (worksheet == null)
            {
                throw new ArgumentException($"The external worksheet '{name}' does not exist.");
            }
            return worksheet;
        }

        /// <summary>
        /// Determines whether the passed value is a absolute external workbook URI
        /// </summary>
        /// <param name="value">Value to check</param>
        /// <returns>True if absolute, otherwise false</returns>
        private static bool IsAbsoluteWorkbookLocation(string value)
        {
            if (IsWindowsDrivePath(value) || IsUncPath(value) || IsUnixRootedPath(value))
            {
                return true;
            }
            return Uri.TryCreate(value, UriKind.Absolute, out Uri parsedUri) && parsedUri.IsAbsoluteUri;
        }

        /// <summary>
        /// Determines whether the passed URI points to a Windows drive (letter)
        /// </summary>
        /// <param name="value">Value to check</param>
        /// <returns>True if a Windows drive path, otherwise false</returns>
        private static bool IsWindowsDrivePath(string value)
        {
            return value.Length >= 3
                && char.IsLetter(value[0])
                && value[1] == ':'
                && (value[2] == '\\' || value[2] == '/');
        }

        /// <summary>
        /// Determines whether the passed URI is a UNC path
        /// </summary>
        /// <param name="value">Value to check</param>
        /// <returns>True if a UNC path, otherwise false</returns>
        private static bool IsUncPath(string value)
        {
            return value.StartsWith(@"\\", StringComparison.Ordinal)
                || value.StartsWith("//", StringComparison.Ordinal);
        }

        /// <summary>
        /// Determines whether the passed URI is a Linux/UNIX drive path
        /// </summary>
        /// <param name="value">Value to check</param>
        /// <returns>True if a Linux/UNIX path, otherwise false</returns>
        private static bool IsUnixRootedPath(string value)
        {
            return value.StartsWith("/", StringComparison.Ordinal) && !value.StartsWith("//", StringComparison.Ordinal);
        }

    }

    /// <summary>
    /// Identifies the OOXML role of an external-link URI relationship.
    /// </summary>
    internal enum ExternalLinkUriRole
    {
        Target,
        AbsoluteAlternate,
        RelativeAlternate
    }

    /// <summary>
    /// Defines one external-link URI relationship for package writing.
    /// </summary>
    internal sealed class ExternalLinkUriRelationship
    {
        public string Id { get; private set; }
        public string RawTarget { get; private set; }
        public string SerializedTarget { get; private set; }
        public ExternalLinkUriRole Role { get; private set; }

        public ExternalLinkUriRelationship(string id, string rawTarget, string serializedTarget, ExternalLinkUriRole role)
        {
            Id = id;
            RawTarget = rawTarget;
            SerializedTarget = serializedTarget;
            Role = role;
        }
    }
}

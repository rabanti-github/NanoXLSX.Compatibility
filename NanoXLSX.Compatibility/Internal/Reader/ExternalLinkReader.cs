/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Reader
{
    /// <summary>
    /// Class implementing a reader for external link files of XLSX files.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_READER", QueueUUID = PlugInUUID.ReaderPrependingQueue, PlugInOrder = 20000)]
    internal class ExternalLinkReader : IDiscoveryPackageReader
    {
        #region privateFields

        private Stream stream;

        private const string ExternalLinkPathRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        private const string OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string AlternateUrlNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021";

        #endregion

        #region properties
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reference to a ReaderPlugInHandler, to be used for inline operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

        public string StreamEntryName => null; // Not used

        public RelationshipInfo CurrentRelationship { get; set; }

        /// <summary>
        /// Document type of external link relationships (in .rels file)
        /// </summary>
        public string DocumentType { get { return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"; } }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public ExternalLinkReader()
        {
        }
        #endregion

        #region methods
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">Stream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Inline plug-in handler</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            Dictionary<int, ExternalWorksheet> worksheets = new Dictionary<int, ExternalWorksheet>();
            string targetRelationshipId = null;
            string absoluteAlternateRelationshipId = null;
            string relativeAlternateRelationshipId = null;
            try
            {
                using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                {
                    bool isExternalBook = false;
                    while (reader.Read())
                    {
                        if (XmlStreamUtils.IsElement(reader, "externalBook"))
                        {
                            isExternalBook = true;
                            targetRelationshipId = reader.GetAttribute("id", OfficeRelationshipNamespace);
                        }

                        if (isExternalBook)
                        {
                            if (reader.NamespaceURI == AlternateUrlNamespace && XmlStreamUtils.IsElement(reader, "absoluteUrl"))
                            {
                                absoluteAlternateRelationshipId = reader.GetAttribute("id", OfficeRelationshipNamespace);
                            }
                            else if (reader.NamespaceURI == AlternateUrlNamespace && XmlStreamUtils.IsElement(reader, "relativeUrl"))
                            {
                                relativeAlternateRelationshipId = reader.GetAttribute("id", OfficeRelationshipNamespace);
                            }
                            else if (XmlStreamUtils.IsElement(reader, "sheetNames"))
                            {
                                GetSheeetNames(reader.ReadSubtree(), worksheets);
                                if (worksheets.Count == 0)
                                {
                                    throw new IOException("No cached worksheets could be determined");
                                }
                            }
                            else if (XmlStreamUtils.IsElement(reader, "sheetDataSet"))
                            {
                                GetSheetData(reader.ReadSubtree(), worksheets);
                            }
                        }
                    }
                }
                ExternalLink link = new ExternalLink();
                RelationshipCatalog discoveryCatalog = Workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                if (discoveryCatalog == null)
                {
                    throw new IOException("The relationship catalog is not available for the external link.");
                }
                string sourcePartPath = CurrentRelationship.ResolvedTargetPath;
                RelationshipInfo target = GetExternalLinkPathRelationship(discoveryCatalog, sourcePartPath, targetRelationshipId, "target", true);
                RelationshipInfo absoluteAlternate = GetExternalLinkPathRelationship(discoveryCatalog, sourcePartPath, absoluteAlternateRelationshipId, "absolute alternate", false);
                RelationshipInfo relativeAlternate = GetExternalLinkPathRelationship(discoveryCatalog, sourcePartPath, relativeAlternateRelationshipId, "relative alternate", false);

                link.SetReadUris(target.Target, absoluteAlternate?.Target, relativeAlternate?.Target);
                foreach (KeyValuePair<int, ExternalWorksheet> worksheet in worksheets)
                {
                    link.AddWorksheet(worksheet.Value);
                }
                link.WorkbookRId = CurrentRelationship.Id;

                List<ExternalLink> externalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
                int index = 0;
                if (externalLinks != null && externalLinks.Count > 0)
                {
                    index = externalLinks.Count;
                }
                Workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY, index, link, true);
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }

        private static RelationshipInfo GetExternalLinkPathRelationship(
            RelationshipCatalog catalog,
            string sourcePartPath,
            string relationshipId,
            string role,
            bool required)
        {
            if (string.IsNullOrEmpty(relationshipId))
            {
                if (required)
                {
                    throw new IOException("The external-link " + role + " relationship ID is missing.");
                }
                return null;
            }

            RelationshipInfo relationship = catalog.GetBySourceAndId(sourcePartPath, relationshipId);
            if (relationship == null)
            {
                throw new IOException("The external-link " + role + " relationship '" + relationshipId + "' could not be resolved.");
            }
            if (!string.Equals(relationship.Type, ExternalLinkPathRelationshipType, StringComparison.Ordinal)
                || relationship.TargetMode != System.IO.Packaging.TargetMode.External)
            {
                throw new IOException("The external-link " + role + " relationship '" + relationshipId + "' has an invalid type or target mode.");
            }
            return relationship;
        }

        private static void GetSheeetNames(XmlReader sheetNames, Dictionary<int, ExternalWorksheet> worksheets)
        {
            int index = 0;
            while (sheetNames.Read())
            {
                if (XmlStreamUtils.IsElement(sheetNames, "sheetName"))
                {
                    if (!string.IsNullOrEmpty(sheetNames.Name))
                    {
                        string val = sheetNames.GetAttribute("val");
                        ExternalWorksheet worksheet = new ExternalWorksheet(val);
                        worksheets[index] = worksheet;
                        index++;
                    }
                }
            }
        }

        private static void GetSheetData(XmlReader sheetDataSet, Dictionary<int, ExternalWorksheet> worksheets)
        {
            int currentIndex = -1;
            while (sheetDataSet.Read())
            {
                if (XmlStreamUtils.IsElement(sheetDataSet, "sheetData"))
                {
                    string id = sheetDataSet.GetAttribute("sheetId");
                    currentIndex = ParserUtils.ParseInt(id);
                    string refreshErrors = sheetDataSet.GetAttribute("refreshErrors");
                    if (refreshErrors != null)
                    {
                        int parserdSate = ParserUtils.ParseBinaryBool(refreshErrors);
                        worksheets[currentIndex].RefreshErros = parserdSate == 1 ? true : false;
                    }
                    continue;
                }
                else if (XmlStreamUtils.IsElement(sheetDataSet, "row"))
                {
                    GetRowData(sheetDataSet.ReadSubtree(), worksheets[currentIndex]);
                }
            }
        }

        private static void GetRowData(XmlReader row, ExternalWorksheet worksheet)
        {
            while (row.Read())
            {
                bool hasCell = false;
                string address = null;
                string type = null;
                string cellMetaData = null; // Roundtrip only
                if (XmlStreamUtils.IsElement(row, "cell"))
                {
                    address = row.GetAttribute("r");
                    type = row.GetAttribute("t");
                    cellMetaData = row.GetAttribute("vm");
                    hasCell = true;
                    continue;
                }
                else if (hasCell && XmlStreamUtils.IsElement(row, "v"))
                {
                    ExternalCellValue.DataType dataType;
                    dataType = ExternalCellValue.DataType.Number;
                    if (type != null)
                    {
                        switch (type)
                        {
                            case "b":
                                dataType = ExternalCellValue.DataType.Boolean;
                                break;
                            case "d":
                                dataType = ExternalCellValue.DataType.Date;
                                break;
                            case "e":
                                dataType = ExternalCellValue.DataType.Error;
                                break;
                            case "s": // Should not be used
                            case "str":
                                dataType = ExternalCellValue.DataType.String;
                                break;
                            default:
                                break;
                        }
                    }
                    string value = row.ReadInnerXml();
                    worksheet.AddCell(address, value, dataType, cellMetaData);
                }
            }
        }
        #endregion
    }

}

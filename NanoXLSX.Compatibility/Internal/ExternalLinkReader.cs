using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class implementing a reader for external link files of XLSX files.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_READER", QueueUUID = PlugInUUID.ReaderPrependingQueue, PlugInOrder = 20000)]
    internal class ExternalLinkReader : IDiscoveryPackageReader
    {
        #region privateFields

        private readonly List<ExternalLink> externalLinks;
        private Stream stream;

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
        /// Reference to a ReaderPlugInHandler, to be used for prepending operations in the <see cref="Execute"/> method
        /// </summary>
        /// 
        /// Reference to a ReaderPlugInHandler, to be used for post operations in the <see cref="Execute"/> method
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
            externalLinks = new List<ExternalLink>();
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
                        }
                        if (isExternalBook && XmlStreamUtils.IsElement(reader, "sheetNames"))
                        {
                            GetSheeetNames(reader.ReadSubtree(), worksheets);
                            if (worksheets.Count == 0)
                            {
                                throw new IOException("No cached worksheets could be determined");
                            }
                        }
                        else if (isExternalBook && XmlStreamUtils.IsElement(reader, "sheetDataSet"))
                        {
                            GetSheetData(reader.ReadSubtree(), worksheets);
                        }
                    }
                }
                ExternalLink link = new ExternalLink();
                RelationshipCatalog discoveryCatalog = Workbook.AuxiliaryData.GetData<RelationshipCatalog>(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity);
                IReadOnlyList<RelationshipInfo> targets = discoveryCatalog
                    .GetByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath")
                    .Where(t => t.SourcePartPath == CurrentRelationship.ResolvedTargetPath)
                    .ToList();
                foreach (RelationshipInfo target in targets)
                {
                    link.AddUri(target.Target);
                }
                foreach (KeyValuePair<int, ExternalWorksheet> worksheet in worksheets)
                {
                    link.AddWorksheet(worksheet.Value);
                }
                link.WorkbookRId = CurrentRelationship.Id;

                List<ExternalLink> externalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
                int index = externalLinks == null ? 0 : externalLinks.Count;
                externalLinks.Add(link);
                Workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY, index, link, true);
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
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
                    continue;
                }
                else if (XmlStreamUtils.IsElement(sheetDataSet, "row"))
                {
                    ExternalWorksheet worksheet = worksheets[currentIndex];
                    GetRowData(sheetDataSet.ReadSubtree(), worksheet);
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
                if (XmlStreamUtils.IsElement(row, "cell"))
                {
                    address = row.GetAttribute("r");
                    type = row.GetAttribute("t");
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
                            case "s":
                                dataType = ExternalCellValue.DataType.SharedString;
                                break;
                            case "str":
                                dataType = ExternalCellValue.DataType.Formula;
                                break;
                            case "inlineStr":
                                dataType = ExternalCellValue.DataType.InlineString;
                                break;
                            default:
                                break;
                        }
                    }
                    string value = row.ReadInnerXml();
                    worksheet.AddCell(address, value, dataType);
                }
            }
        }

        #endregion

        #region sub-classes

        #endregion
    }

}
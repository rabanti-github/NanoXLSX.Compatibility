using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace NanoXLSX.Internal.Writers
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_PACKAGE_WRITER", QueueUUID = PlugInUUID.WriterPackageRegistryQueue, PlugInOrder = 20000)]
    internal class ExternalLinkPackageWriter : IPluginPackageWriter
    {

        private const string packagePartPath = "xl/externalLinks/";
        private const string contentType = @"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        private const string relationshipType = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        private int currentOrderNr;

        public XmlElement XmlElement => null; // NoOp in this plug-in type

        public Workbook Workbook { get; set; }

        public int CurrentIndex { get; set; }

        public List<int> OrderNumbers { get; private set; }

        public List<string> PackagePartPaths { get; private set; }

        public List<string> PackagePartFileNames { get; private set; }

        public List<string> ContentTypes { get; private set; }

        public List<string> RelationshipTypes { get; private set; }

        public List<bool> ArePackagePartsRoot { get; private set; }

        public List<XmlElement> XmlElements { get; private set; }

        public ExternalLinkPackageWriter()
        {
            CurrentIndex = -1;
            OrderNumbers = new List<int>();
            PackagePartPaths = new List<string>();
            PackagePartFileNames = new List<string>();
            ContentTypes = new List<string>();
            RelationshipTypes = new List<string>();
            ArePackagePartsRoot = new List<bool>();
            XmlElements = new List<XmlElement>();
        }

        public void Init(IBaseWriter baseWriter)
        {
            Workbook = baseWriter.Workbook;
            int nr = Workbook.AuxiliaryData.GetData<int>(PlugInUUID.WriterPackageRegistryQueue, PlugInUUID.LastPackageOrderNumber);
            List<ExternalLink> externalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
            currentOrderNr = nr + 1000;

            if (externalLinks == null || externalLinks.Count == 0)
            {
                return; // Nothing to register and write
            }
            for (int i = 0; i < externalLinks.Count; i++)
            {
                ExternalLink externalLink = externalLinks[i];
                currentOrderNr++;
                string name = "externalLink" + ParserUtils.ToString(i + 1) + ".xml";
                OrderNumbers.Add(currentOrderNr);
                PackagePartPaths.Add(packagePartPath);
                PackagePartFileNames.Add(name);
                ContentTypes.Add(contentType);
                RelationshipTypes.Add(relationshipType);
                ArePackagePartsRoot.Add(false);
                XmlElements.Add(GetElement(externalLink));
            }
        }

        public void Execute()
        {
            // NoOp
        }

        internal static XmlElement GetElement(ExternalLink externalLink)
        {
            IReadOnlyList<ExternalLinkUriRelationship> relationships = externalLink.GetUriRelationships();
            XmlElement element = XmlElement.CreateElement("externalLink");
            element.AddDefaultXmlNameSpace("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            element.AddNameSpaceAttribute("mc", "xmlns", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            element.AddNameSpaceAttribute("x14", "xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            element.AddNameSpaceAttribute("xxl21", "xmlns", "http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021");
            element.AddAttribute("mc:Ignorable", "x14 xxl21");

            XmlElement externalBook = XmlElement.CreateElement("externalBook");
            externalBook.AddNameSpaceAttribute("r", "xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            externalBook.AddAttribute("r:id", relationships[0].Id); // There is one primary external book relationship per external-link part.
            element.AddChildElement(externalBook);

            if (relationships.Count > 1)
            {
                XmlElement alternateUrls = XmlElement.CreateElement("alternateUrls", "xxl21");
                foreach (ExternalLinkUriRelationship relationship in relationships)
                {
                    if (relationship.Role == ExternalLinkUriRole.AbsoluteAlternate)
                    {
                        XmlElement absoluteUrl = XmlElement.CreateElement("absoluteUrl", "xxl21");
                        absoluteUrl.AddAttribute("r:id", relationship.Id);
                        alternateUrls.AddChildElement(absoluteUrl);
                    }
                    else if (relationship.Role == ExternalLinkUriRole.RelativeAlternate)
                    {
                        XmlElement relativeUrl = XmlElement.CreateElement("relativeUrl", "xxl21");
                        relativeUrl.AddAttribute("r:id", relationship.Id);
                        alternateUrls.AddChildElement(relativeUrl);
                    }
                }
                externalBook.AddChildElement(alternateUrls);
            }
            if (externalLink.Worksheets.Count > 0)
            {
                XmlElement sheetNames = XmlElement.CreateElement("sheetNames");
                XmlElement sheetDataSet = XmlElement.CreateElement("sheetDataSet");
                // Despite ISO/IEC 29500 describing sheetId as a 1-based index, Microsoft Excel writes it as a zero-based index into sheetNames.
                int sheetId = 0;
                foreach (ExternalWorksheet sheet in externalLink.Worksheets)
                {
                    // --- Sheet names
                    XmlElement sheetName = XmlElement.CreateElement("sheetName");
                    sheetName.AddAttribute("val", XmlUtils.SanitizeXmlValue(sheet.Name));
                    sheetNames.AddChildElement(sheetName);
                    // --- Sheet data
                    XmlElement sheetData = XmlElement.CreateElement("sheetData");
                    sheetData.AddAttribute("sheetId", ParserUtils.ToString(sheetId));
                    if (sheet.RefreshErros != null)
                    {
                        // currently only for roundtrip
                        sheetData.AddAttribute("refreshErrors", ParserUtils.ToString(sheet.RefreshErros.Value == true ? 1 : 0));
                    }
                    List<XmlElement> row = GetRowData(sheet);
                    if (row.Count > 0)
                    {
                        foreach (XmlElement rowElement in row)
                        {
                            sheetData.AddChildElement(rowElement);
                        }
                    }
                    sheetDataSet.AddChildElement(sheetData);
                    sheetId++;
                }
                externalBook.AddChildElement(sheetNames);
                externalBook.AddChildElement(sheetDataSet);
            }
            if (externalLink.DefinedNames.Count > 0)
            {
                XmlElement definedNames = XmlElement.CreateElement("definedNames");
                foreach (ExternalDefinedName externalDefinedName in externalLink.DefinedNames)
                {
                    XmlElement definedName = XmlElement.CreateElement("definedName");
                    definedName.AddAttribute("name", XmlUtils.SanitizeXmlValue(externalDefinedName.Name));
                    if (externalDefinedName.RefersTo != null)
                    {
                        definedName.AddAttribute("refersTo", XmlUtils.SanitizeXmlValue(externalDefinedName.RefersTo));
                    }
                    if (externalDefinedName.RelationshipId != null) // currently only for roundtrip
                    {
                        definedName.AddAttribute("sheetId", XmlUtils.SanitizeXmlValue(externalDefinedName.RelationshipId)); // sanitized (could be unsafe)
                    }
                    definedNames.AddChildElement(definedName);
                }
                externalBook.AddChildElement(definedNames);
            }
            return element;
        }

        private static List<XmlElement> GetRowData(ExternalWorksheet sheet)
        {
            if (sheet.Cells.Count == 0)
            {
                return new List<XmlElement>(); ;
            }
            ReadOnlyDictionary<Address, ExternalCellValue> cells = sheet.Cells;
            // List<Address> addresses = cells.Keys.ToList();

            SortedDictionary<int, XmlElement> rows = new SortedDictionary<int, XmlElement>();
            foreach (KeyValuePair<Address, ExternalCellValue> cell in cells)
            {
                if (!rows.TryGetValue(cell.Key.Row, out XmlElement value))
                {
                    XmlElement row = XmlElement.CreateElement("row");
                    row.AddAttribute("r", ParserUtils.ToString(cell.Key.Row + 1)); // 1-based
                    value = row;
                    rows.Add(cell.Key.Row, value);
                }
                XmlElement element = XmlElement.CreateElement("cell");
                element.AddAttribute("r", cell.Key.ToString());
                string type = GetCellType(cell.Value.Type);
                if (type != null)
                {
                    element.AddAttribute("t", type);
                }
                if (cell.Value.CellMetadata != null) // currently only for roundtrip
                {
                    element.AddAttribute("vm", ParserUtils.ToString(cell.Value.CellMetadata.Value));
                }
                if (!string.IsNullOrEmpty(cell.Value.Value))
                {
                    element.InnerValue = XmlUtils.SanitizeXmlValue(cell.Value.Value);
                }
                value.AddChildElement(element);
            }
            return rows.Values.ToList();
        }

        private static string GetCellType(ExternalCellValue.DataType dataType)
        {
            switch (dataType)
            {
                case ExternalCellValue.DataType.Boolean:
                    return "b";
                case ExternalCellValue.DataType.Date:
                    return "d";
                case ExternalCellValue.DataType.Error:
                    return "e";
                case ExternalCellValue.DataType.SharedString:
                    return "s";
                case ExternalCellValue.DataType.Formula:
                    return "str";
                default:
                    return null; // numeric
            }
        }
    }
}

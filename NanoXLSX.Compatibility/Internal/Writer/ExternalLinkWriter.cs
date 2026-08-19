/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace NanoXLSX.Internal.Writer
{
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_WRITER", QueueUUID = PlugInUUID.WriterAppendingQueue, PlugInOrder = 20001)]
    internal class ExternalLinkWriter : IPluginIndexedWriter
    {
        private string currentUniqueIndex;
        private List<ExternalLink> externalLinks;
        private int maxIndex;
        XmlElement xmlElement;

        public int CurrentIndex { get; set; }

        public string CurrentUniquePackagePartIndex => currentUniqueIndex;

        public int MaxIndex => maxIndex;

        public Workbook Workbook { get; set; }

        public XmlElement XmlElement => xmlElement;

        public void Init(IBaseWriter baseWriter)
        {
            this.Workbook = baseWriter.Workbook;
            List<ExternalLink> storedExternalLinks = Workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
            externalLinks = storedExternalLinks == null ? new List<ExternalLink>() : storedExternalLinks.OfType<ExternalLink>().ToList();
            maxIndex = externalLinks.Count - 1;
        }

        public void Execute()
        {
            currentUniqueIndex = CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + ParserUtils.ToString(CurrentIndex);
            xmlElement = GetElement(externalLinks[CurrentIndex]);
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
            // XSD: sheetNames > definedNames > sheetDataSet
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
                if (externalBook.FindChildElementsByName("sheetDataSet").Any())
                {
                    externalBook.AddChildElementBefore(definedNames, "sheetDataSet");
                }
                else if (externalBook.FindChildElementsByName("sheetNames").Any())
                {
                    externalBook.AddChildElementAfter(definedNames, "sheetNames");
                }
                else
                {
                    externalBook.AddChildElement(definedNames);
                }
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
                if (cell.Value.Type != ExternalCellValue.DataType.Empty)
                {
                    XmlElement valueElement = XmlElement.CreateElement("v");
                    valueElement.InnerValue =
                        XmlUtils.SanitizeXmlValue(cell.Value.Value);
                    element.AddChildElement(valueElement);
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
                case ExternalCellValue.DataType.String: // s is not used
                    return "str";
                default:
                    return null; // numeric
            }
        }

    }
}

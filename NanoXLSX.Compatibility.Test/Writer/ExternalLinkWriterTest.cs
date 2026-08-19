using NanoXLSX.Internal;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Utils.Xml;
using System;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class ExternalLinkWriterTest
    {
        [Fact(DisplayName = "Test initialization and execution for two external links")]
        public void InitializesAndExecutesForTwoExternalLinks()
        {
            Workbook workbook = new Workbook("Sheet1");
            ExternalLink first = new ExternalLink(@"C:\data\first.xlsx");
            ExternalLink second = new ExternalLink(@"C:\data\second.xlsx");
            ExternalLinkTestUtils.StoreExternalLink(workbook, first, 0);
            ExternalLinkTestUtils.StoreExternalLink(workbook, second, 1);
            ExternalLinkWriter writer = new ExternalLinkWriter();

            writer.Init(new ExternalLinkTestUtils.TestBaseWriter(workbook));
            writer.CurrentIndex = 1;
            writer.Execute();

            Assert.Equal(1, writer.MaxIndex); // 1 = 2 elements
            Assert.Equal(CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + "1", writer.CurrentUniquePackagePartIndex);
            Assert.Equal("externalLink", writer.XmlElement.Name);
            XmlElement externalBook = Assert.Single(writer.XmlElement.FindChildElementsByName("externalBook"));
            Assert.Equal("rId1", ExternalLinkTestUtils.GetAttribute(externalBook, "r:id"));
        }

        [Fact(DisplayName = "Test failure when executing the external-link writer without a valid index")]
        public void RejectsInvalidExternalLinkIndex()
        {
            ExternalLinkWriter writer = new ExternalLinkWriter();
            writer.Init(new ExternalLinkTestUtils.TestBaseWriter(new Workbook("Sheet1")));
            writer.CurrentIndex = 0;

            Assert.Throws<ArgumentOutOfRangeException>(() => writer.Execute());
            Assert.Equal(-1, writer.MaxIndex);
        }

        [Fact(DisplayName = "Test writing alternate external-link URI relationships")]
        public void WritesAlternateUriRelationships()
        {
            ExternalLink link = new ExternalLink();
            link.SetReadUris(@"..\data\external.xlsx", @"C:\data\external.xlsx", @"alternate\external.xlsx");

            XmlElement element = ExternalLinkWriter.GetElement(link);

            XmlElement externalBook = Assert.Single(element.FindChildElementsByName("externalBook"));
            XmlElement alternateUrls = Assert.Single(externalBook.FindChildElementsByName("alternateUrls"));
            XmlElement absolute = Assert.Single(alternateUrls.FindChildElementsByName("absoluteUrl"));
            XmlElement relative = Assert.Single(alternateUrls.FindChildElementsByName("relativeUrl"));
            Assert.Equal("rId2", ExternalLinkTestUtils.GetAttribute(absolute, "r:id"));
            Assert.Equal("rId3", ExternalLinkTestUtils.GetAttribute(relative, "r:id"));
        }

        [Fact(DisplayName = "Test writing sorted cached worksheet rows, metadata, and empty cells")]
        public void WritesCachedWorksheetRowsAndMetadata()
        {
            ExternalLink link = new ExternalLink(@"C:\data\external.xlsx");
            ExternalWorksheet emptySheet = new ExternalWorksheet("Empty");
            emptySheet.RefreshErros = false;
            ExternalWorksheet dataSheet = new ExternalWorksheet("Data");
            dataSheet.RefreshErros = true;
            dataSheet.AddCell("B3", "later", ExternalCellValue.DataType.String);
            dataSheet.AddCell("A1", null, ExternalCellValue.DataType.Empty);
            dataSheet.AddCell("C1", "42", ExternalCellValue.DataType.Number);
            dataSheet.Cells[new Address("C1")].CellMetadata = 7;
            link.AddWorksheet(emptySheet);
            link.AddWorksheet(dataSheet);

            XmlElement element = ExternalLinkWriter.GetElement(link);

            XmlElement externalBook = Assert.Single(element.FindChildElementsByName("externalBook"));
            XmlElement sheetNames = Assert.Single(externalBook.FindChildElementsByName("sheetNames"));
            Assert.Equal(new[] { "Empty", "Data" }, sheetNames.Children.ConvertAll(item => ExternalLinkTestUtils.GetAttribute(item, "val")));
            XmlElement dataSet = Assert.Single(externalBook.FindChildElementsByName("sheetDataSet"));
            Assert.Equal("0", ExternalLinkTestUtils.GetAttribute(dataSet.Children[0], "refreshErrors"));
            Assert.Equal("1", ExternalLinkTestUtils.GetAttribute(dataSet.Children[1], "refreshErrors"));
            Assert.Null(dataSet.Children[0].Children);
            Assert.Equal(new[] { "1", "3" }, dataSet.Children[1].Children.ConvertAll(row => ExternalLinkTestUtils.GetAttribute(row, "r")));
            XmlElement emptyCell = Assert.Single(dataSet.Children[1].FindChildElementsByNameAndAttribute("cell", "r", "A1"));
            Assert.Empty(emptyCell.FindChildElementsByName("v"));
            XmlElement metadataCell = Assert.Single(dataSet.Children[1].FindChildElementsByNameAndAttribute("cell", "r", "C1"));
            Assert.Equal("7", ExternalLinkTestUtils.GetAttribute(metadataCell, "vm"));
        }

        [Theory(DisplayName = "Test writing cached external-cell data types")]
        [InlineData(ExternalCellValue.DataType.Number, "42", null)]
        [InlineData(ExternalCellValue.DataType.Boolean, "1", "b")]
        [InlineData(ExternalCellValue.DataType.Date, "2026-08-19", "d")]
        [InlineData(ExternalCellValue.DataType.Error, "#N/A", "e")]
        [InlineData(ExternalCellValue.DataType.String, "text", "str")]
        [InlineData((ExternalCellValue.DataType)999, "custom", null)]
        public void WritesCachedCellType(ExternalCellValue.DataType dataType, string value, string expectedType)
        {
            ExternalLink link = new ExternalLink(@"C:\data\external.xlsx");
            ExternalWorksheet sheet = new ExternalWorksheet("Data");
            sheet.AddCell("A1", value, dataType);
            link.AddWorksheet(sheet);

            XmlElement element = ExternalLinkWriter.GetElement(link);

            XmlElement cell = Assert.Single(element.FindChildElementsByName("cell"));
            Assert.Equal(expectedType, ExternalLinkTestUtils.GetAttribute(cell, "t"));
            Assert.Equal(value, Assert.Single(cell.FindChildElementsByName("v")).InnerValue);
        }

        [Theory(DisplayName = "Test writing external defined names with and without worksheet data")]
        [InlineData(false)]
        [InlineData(true)]
        public void WritesExternalDefinedNamesInSchemaOrder(bool includeWorksheet)
        {
            ExternalLink link = new ExternalLink(@"C:\data\external.xlsx");
            ExternalDefinedName withReference = new ExternalDefinedName("WithReference", "Data!$A$1");
            withReference.RelationshipId = "2";
            ExternalDefinedName withoutReference = new ExternalDefinedName("WithoutReference", null, false);
            link.AddDefinedName(withReference);
            link.AddDefinedName(withoutReference);
            if (includeWorksheet)
            {
                link.AddWorksheet(new ExternalWorksheet("Data"));
            }

            XmlElement element = ExternalLinkWriter.GetElement(link);

            XmlElement externalBook = Assert.Single(element.FindChildElementsByName("externalBook"));
            XmlElement definedNames = Assert.Single(externalBook.FindChildElementsByName("definedNames"));
            Assert.Equal("Data!$A$1", ExternalLinkTestUtils.GetAttribute(definedNames.Children[0], "refersTo"));
            Assert.Equal("2", ExternalLinkTestUtils.GetAttribute(definedNames.Children[0], "sheetId"));
            Assert.Null(ExternalLinkTestUtils.GetAttribute(definedNames.Children[1], "refersTo"));
            if (includeWorksheet)
            {
                Assert.Equal(new[] { "sheetNames", "definedNames", "sheetDataSet" }, externalBook.Children.GetRange(0, 3).ConvertAll(child => child.Name));
            }
            else
            {
                Assert.Equal("definedNames", externalBook.Children[0].Name);
            }
        }
    }
}

using NanoXLSX.Internal;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;
using System;
using System.IO;
using System.IO.Packaging;
using System.Text;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalLinkTest
    {
        [Fact(DisplayName = "Test of the properties")]
        public void PropertiesTest()
        {
            ExternalLink link = new ExternalLink(@"C:\Files\external.xlsx", @"..\external.xlsx");
            ExternalWorksheet ws = new ExternalWorksheet("extName");

            link.AddDefinedName("name", "link");
            link.AddWorksheet(ws);
            link.WorkbookRId = "rId5";

            Assert.Single(link.DefinedNames);
            Assert.Single(link.Worksheets);

            Assert.Equal(@"C:\Files\external.xlsx", link.AbsoluteUri);
            Assert.Equal(@"..\external.xlsx", link.RelativeUri);
            Assert.Equal(@"..\external.xlsx", link.TargetUri);
            Assert.Equal(@"C:\Files\external.xlsx", link.AbsoluteAlternateUri);
            Assert.Null(link.RelativeAlternateUri);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("extName", link.Worksheets[0].Name);
            Assert.Equal("rId5", link.WorkbookRId);
            Assert.Equal("[C:\\Files\\external.xlsx]", link.ReadableReferenceToken);
        }

        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.NotNull(link.DefinedNames);
            Assert.NotNull(link.Worksheets);
            Assert.Empty(link.DefinedNames);
            Assert.Empty(link.Worksheets);
            Assert.Null(link.AbsoluteUri);
            Assert.Null(link.RelativeUri);
            Assert.Null(link.TargetUri);
            Assert.Null(link.WorkbookRId);
            Assert.Null(link.ReadableReferenceToken);
        }

        [Fact(DisplayName = "Test of the absolute URI constructor")]
        public void ConstructorTest2()
        {
            ExternalLink link = new ExternalLink(@"C:\Files\external.xlsx");
            Assert.Empty(link.DefinedNames);
            Assert.Empty(link.Worksheets);
            Assert.Equal(@"C:\Files\external.xlsx", link.AbsoluteUri);
            Assert.Null(link.RelativeUri);
            Assert.Equal(link.AbsoluteUri, link.TargetUri);
            Assert.Null(link.AbsoluteAlternateUri);
            Assert.Null(link.RelativeAlternateUri);
        }

        [Theory(DisplayName = "Test of accepted absolute workbook locations")]
        [InlineData(@"C:\Files\external.xlsx")]
        [InlineData(@"C:/Files/external.anything")]
        [InlineData(@"\\server\share\external.xlsm")]
        [InlineData(@"/mnt/files/external.xls")]
        [InlineData(@"file:///C:/Files/external.xlsx")]
        [InlineData(@"https://example.com/files/external.xlsx")]
        public void AbsoluteWorkbookLocationsTest(string uri)
        {
            ExternalLink link = new ExternalLink(uri);
            Assert.Equal(uri, link.AbsoluteUri);
            Assert.Equal(uri, link.TargetUri);
        }

        [Theory(DisplayName = "Test of invalid absolute workbook locations")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("external.xlsx")]
        [InlineData("NoExtensionFile")]
        [InlineData(@"C:\Files\external")]
        [InlineData(@"..\Files\external\file.xlsx")]
        [InlineData(@"C:\Files\")]
        [InlineData(@"file://server")]
        [InlineData(@"https://example.com/files/")]
        public void FailingConstructorTest(string uri)
        {
            Assert.Throws<ArgumentException>(() => new ExternalLink(uri));
        }

        [Theory(DisplayName = "Test of accepted relative workbook locations")]
        [InlineData("external.xlsx")]
        [InlineData(@"..\files\external.xlsx")]
        [InlineData(@"folder/external.custom")]
        public void RelativeWorkbookLocationsTest(string uri)
        {
            ExternalLink link = new ExternalLink(@"C:\Files\external.xlsx", uri);
            Assert.Equal(uri, link.RelativeUri);
            Assert.Equal(uri, link.TargetUri);
            Assert.Equal(link.AbsoluteUri, link.AbsoluteAlternateUri);
        }

        [Theory(DisplayName = "Test of invalid relative workbook locations")]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("external")]
        [InlineData(@"folder\")]
        [InlineData(@"C:\Files\external.xlsx")]
        [InlineData(@"\\server\share\external.xlsx")]
        [InlineData(@"/mnt/files/external.xlsx")]
        [InlineData(@"https://example.com/external.xlsx")]
        public void FailingRelativeWorkbookLocationsTest(string uri)
        {
            Assert.Throws<ArgumentException>(() => new ExternalLink(@"C:\Files\external.xlsx", uri));
        }

        [Theory(DisplayName = "Test of the ReadableReferenceToken property")]
        [InlineData(@"C:\d\wb.xlsx", null, null, @"[C:\d\wb.xlsx]")]
        [InlineData(@"C:\d\wb.xlsx", @"C:\d\wb2.xlsx", null, @"[C:\d\wb.xlsx]")]
        [InlineData(@"C:\d\wb.xlsx", @"C:\d\wb2.xlsx", @"C:\d\wb3.xlsx", @"[C:\d\wb.xlsx]")]
        [InlineData(@"..\f\wb1.xlsx", @"C:\d\wb2.xlsx", null, @"[..\f\wb1.xlsx]")]
        [InlineData(@"..\f\wb1.xlsx", @"..\f\wb2.xlsx", @"[C:\d\wb.xlsx]", @"[..\f\wb1.xlsx]")]
        public void ReadableReferenceTokenTest(string targetUri, string absoluteAlternateUri, string relativeAlternateUri, string expectedToken)
        {
            ExternalLink link = new ExternalLink();
            link.SetReadUris(targetUri, null, null);
            Assert.Equal(expectedToken, link.ReadableReferenceToken);
        }

        [Fact(DisplayName = "Test of URI roles populated by the reader")]
        public void ReadUrisTest()
        {
            ExternalLink link = new ExternalLink();
            link.SetReadUris(@"..\files\external.xlsx", @"file:///C:/Files/external.xlsx", @"alternative/external.xlsx");

            Assert.Equal(@"file:///C:/Files/external.xlsx", link.AbsoluteUri);
            Assert.Equal(@"..\files\external.xlsx", link.RelativeUri);
            Assert.Equal("[file:///C:/Files/external.xlsx]", link.ReadableReferenceToken);
            Assert.Equal(3, link.GetWorkbookLocations().Count);

            var relationships = link.GetUriRelationships();
            Assert.Equal("rId1", relationships[0].Id);
            Assert.Equal(ExternalLinkUriRole.Target, relationships[0].Role);
            Assert.Equal("rId2", relationships[1].Id);
            Assert.Equal(ExternalLinkUriRole.AbsoluteAlternate, relationships[1].Role);
            Assert.Equal("file:///C:/Files/external.xlsx", relationships[1].SerializedTarget);
            Assert.Equal("rId3", relationships[2].Id);
            Assert.Equal(ExternalLinkUriRole.RelativeAlternate, relationships[2].Role);
            Assert.Equal("alternative/external.xlsx", relationships[2].SerializedTarget);
        }

        [Fact(DisplayName = "Test of duplicate URI roles populated by the reader")]
        public void DuplicateReadUrisTest()
        {
            const string uri = @"file:///C:/Files/external.xlsx";
            ExternalLink link = new ExternalLink();
            link.SetReadUris(uri, uri, null);

            Assert.Single(link.GetWorkbookLocations());
            Assert.Equal(2, link.GetUriRelationships().Count);
            Assert.Equal(uri, link.AbsoluteUri);
        }

        [Theory(DisplayName = "Test of relationship target serialization")]
        [InlineData(@"C:\Documents\Codex Quota.xlsx", "file:///C:/Documents/Codex%20Quota.xlsx")]
        [InlineData(@"\\server\share\CodexQuota.xlsx", "file://server/share/CodexQuota.xlsx")]
        [InlineData(@"/mnt/files/Codex Quota.xlsx", "file:///mnt/files/Codex%20Quota.xlsx")]
        [InlineData(@"file:///C:\Documents\Codex Quota.xlsx", "file:///C:/Documents/Codex%20Quota.xlsx")]
        [InlineData(@"file:///C:/Documents/Codex%20Quota.xlsx", "file:///C:/Documents/Codex%20Quota.xlsx")]
        [InlineData(@"https://example.com/Codex%20Quota.xlsx", "https://example.com/Codex%20Quota.xlsx")]
        [InlineData(@"..\files\Codex Quota.xlsx", "../files/Codex%20Quota.xlsx")]
        [InlineData(@"../files/Codex%20Quota.xlsx?version=1#sheet", "../files/Codex%20Quota.xlsx?version=1#sheet")]
        [InlineData(@"/file.xlsx", "file:///file.xlsx")]
        [InlineData("../files/book.xlsx?version=1", "../files/book.xlsx?version=1")]
        [InlineData("../files/book.xlsx#Sheet1", "../files/book.xlsx#Sheet1")]
        [InlineData("https://example.com/files/book.xlsx?version=1#Sheet1", "https://example.com/files/book.xlsx?version=1#Sheet1")]
        public void SerializeRelationshipTargetTest(string input, string expected)
        {
            Assert.Equal(expected, ExternalLink.SerializeRelationshipTarget(input));
        }

        [Fact(DisplayName = "Test of alternate URL package XML")]
        public void ExternalLinkPackageXmlTest()
        {
            ExternalLink link = new ExternalLink();
            link.SetReadUris(@"..\files\external.xlsx", @"C:\Files\external.xlsx", @"alternative\external.xlsx");

            XmlElement root = ExternalLinkPackageWriter.GetElement(link);
            XmlElement externalBook = Assert.Single(root.Children);
            Assert.Equal("externalBook", externalBook.Name);
            Assert.Equal("rId1", XmlAttribute.FindAttribute("r:id", externalBook.Attributes).Value.Value);

            XmlElement alternateUrls = Assert.Single(externalBook.Children);
            Assert.Equal("alternateUrls", alternateUrls.Name);
            Assert.Equal(2, alternateUrls.Children.Count);
            Assert.Equal("absoluteUrl", alternateUrls.Children[0].Name);
            Assert.Equal("rId2", XmlAttribute.FindAttribute("r:id", alternateUrls.Children[0].Attributes).Value.Value);
            Assert.Equal("relativeUrl", alternateUrls.Children[1].Name);
            Assert.Equal("rId3", XmlAttribute.FindAttribute("r:id", alternateUrls.Children[1].Attributes).Value.Value);
        }

        [Fact(DisplayName = "Test of URI role resolution by the external-link reader")]
        public void ExternalLinkReaderUriRolesTest()
        {
            const string sourcePart = "xl/externalLinks/externalLink1.xml";
            const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
            const string documentType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
            // TODO provide a test file instead
            const string xml =
                "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:xxl21=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\">" +
                "<externalBook r:id=\"rId1\"><xxl21:alternateUrls>" +
                "<xxl21:absoluteUrl r:id=\"rId2\"/><xxl21:relativeUrl r:id=\"rId3\"/>" +
                "</xxl21:alternateUrls></externalBook></externalLink>";

            RelationshipCatalog catalog = new RelationshipCatalog();
            catalog.TryAdd(new RelationshipInfo("rId3", relationshipType, "alternative/external.xlsx", TargetMode.External, null, sourcePart, null));
            catalog.TryAdd(new RelationshipInfo("rId1", relationshipType, @"..\files\external.xlsx", TargetMode.External, null, sourcePart, null));
            catalog.TryAdd(new RelationshipInfo("rId2", relationshipType, "file:///C:/Files/external.xlsx", TargetMode.External, null, sourcePart, null));

            Workbook workbook = new Workbook();
            workbook.AuxiliaryData.SetData(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity, catalog);
            RelationshipInfo currentRelationship = new RelationshipInfo(
                "rId7", documentType, "externalLinks/externalLink1.xml", TargetMode.Internal,
                "xl/_rels/workbook.xml.rels", "xl/workbook.xml", sourcePart);

            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                ExternalLinkReader reader = new ExternalLinkReader
                {
                    CurrentRelationship = currentRelationship
                };
                reader.Init(stream, workbook, null, null);
                reader.Execute();
            }

            ExternalLink result = Assert.Single(workbook.AuxiliaryData.GetDataList<ExternalLink>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY));
            Assert.Equal(@"..\files\external.xlsx", result.TargetUri);
            Assert.Equal("file:///C:/Files/external.xlsx", result.AbsoluteAlternateUri);
            Assert.Equal("alternative/external.xlsx", result.RelativeAlternateUri);
            Assert.Equal("rId7", result.WorkbookRId);
        }

        [Fact(DisplayName = "Test of the AddDefinedName method")]
        public void AddDefinedNameTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            link.AddDefinedName(name);
            Assert.Single(link.DefinedNames);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("refTo", link.DefinedNames[0].RefersTo);
        }

        [Fact(DisplayName = "Test of the AddDefinedName method on multiple values")]
        public void AddMultipleDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            ExternalDefinedName name2 = new ExternalDefinedName("name2", "refTo2");
            link.AddDefinedName(name);
            link.AddDefinedName(name2);
            Assert.Equal(2, link.DefinedNames.Count);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("refTo", link.DefinedNames[0].RefersTo);
            Assert.Equal("name2", link.DefinedNames[1].Name);
            Assert.Equal("refTo2", link.DefinedNames[1].RefersTo);
        }

        [Fact(DisplayName = "Test of the failing AddDefinedName method on null")]
        public void FailingAddDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.Throws<ArgumentException>(() => { link.AddDefinedName(null); });
        }

        [Fact(DisplayName = "Test of the failing AddDefinedName method on duplicate values")]
        public void FailingDuplicateDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            link.AddDefinedName(name);
            ExternalDefinedName name2 = new ExternalDefinedName("name", "refTo2");
            Assert.Throws<ArgumentException>(() => { link.AddDefinedName(name2); });
        }

        [Fact(DisplayName = "Test of the AddWorksheet method")]
        public void AddWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Single(link.Worksheets);
            Assert.Equal("ext", link.Worksheets[0].Name);
        }

        [Fact(DisplayName = "Test of the AddWorksheet method on multiple values")]
        public void AddMultipleWorksheetsTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            ExternalWorksheet ws2 = new ExternalWorksheet("ext2");
            link.AddWorksheet(ws);
            link.AddWorksheet(ws2);
            Assert.Equal(2, link.Worksheets.Count);
            Assert.Equal("ext", link.Worksheets[0].Name);
            Assert.Equal("ext2", link.Worksheets[1].Name);
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet method on null")]
        public void FailingAddWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.Throws<ArgumentException>(() => { link.AddWorksheet(null); });
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet method on duplicate values")]
        public void FailingDuplicateWorksheetsTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            ExternalWorksheet ws2 = new ExternalWorksheet("ext");
            Assert.Throws<ArgumentException>(() => { link.AddWorksheet(ws2); });
        }

        [Fact(DisplayName = "Test of the GetWorksheet method")]
        public void GetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            ExternalWorksheet ws2 = new ExternalWorksheet("ext2");
            link.AddWorksheet(ws);
            link.AddWorksheet(ws2);
            ExternalWorksheet restult = link.GetWorksheet("ext");
            Assert.NotNull(restult);
            Assert.Equal("ext", restult.Name);
        }

        [Fact(DisplayName = "Test of the failing GetWorksheet method on null")]
        public void FailingGetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Throws<ArgumentException>(() => { link.GetWorksheet(null); });
        }

        [Fact(DisplayName = "Test of the failing GetWorksheet on a not existing name")]
        public void FailingNotExistingGetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Throws<ArgumentException>(() => { link.GetWorksheet("ext2"); });
        }

        #region builderTests

        [Fact(DisplayName = "Test of the CreateBuilder method in ExternalLink")]
        public void CreateBuilderTest()
        {
            ExternalLink link = new ExternalLink(@"C:\Files\external.xlsx");
            var builder = link.CreateBuilder();
            Assert.NotNull(builder);
            Assert.IsType<ExternalLinkBuilder>(builder);
        }

        [Fact(DisplayName = "Test of the external link passing function of the builder")]
        public void BuilderAddExternalLinkTest()
        {
            ExternalLink extLink = new ExternalLink(@"C:\Files\external.xlsx");
            ExternalLinkBuilder passed = new ExternalLinkBuilder(extLink);
            ExternalLink result = passed.Build();
            Assert.NotNull(result);
            Assert.Same(extLink, result);
            Assert.Equal(@"C:\Files\external.xlsx", result.AbsoluteUri);
        }

        [Fact(DisplayName = "Test of the failing external link passing function of the builder on null")]
        public void BuilderAddExternalLinkFailTest()
        {
            Assert.Throws<ArgumentException>(() => { var builder = new ExternalLinkBuilder(null); });
        }

        [Fact(DisplayName = "Test of the external link builder function for adding worksheets")]
        public void BuilderWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            ExternalLink result = builder.Build();
            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
        }

        [Theory(DisplayName = "Test of the failing external link builder function for adding worksheets on invalid values")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t")]
        [InlineData("a")] // duplicate
        public void BuilderFailingWorksheetTest(string name)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.ThrowsAny<Exception>(() => { builder.AddWorksheet("a").AddWorksheet(name); });
        }

        [Fact(DisplayName = "Test of the external link builder function for using worksheets")]
        public void BuilderUseWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            builder.UseWorksheet("ws1");
            builder.AddCell("A1", "test");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
            Assert.Empty(result.Worksheets[1].Cells);
            Assert.Single(result.Worksheets[0].Cells);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[0].Cells[new Address("A1")].Value);
        }

        [Fact(DisplayName = "Test of the external link builder function for using worksheets automatically on add")]
        public void BuilderUsingAddedWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            builder.AddCell("A1", "test");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
            Assert.Empty(result.Worksheets[0].Cells);
            Assert.Single(result.Worksheets[1].Cells);
            Assert.NotNull(result.Worksheets[1].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[1].Cells[new Address("A1")].Value);
        }

        [Theory(DisplayName = "Test of the failing external link builder function for using worksheets on a invalid value")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("ws3")]
        public void BuilderUseWorksheetFailTest(string name)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            Assert.ThrowsAny<Exception>(() => { builder.UseWorksheet(name); });
        }

        [Fact(DisplayName = "Test of the external link builder function for adding cells")]
        public void BuilderAddCellTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1");
            builder.AddCell("A1", "test");
            builder.AddCell("B2", "test2");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets[0].Cells.Count);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[0].Cells[new Address("A1")].Value);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("B2")]);
            Assert.Equal("test2", result.Worksheets[0].Cells[new Address("B2")].Value);
        }

        [Fact(DisplayName = "Test of the external link builder function for adding cells (overload)")]
        public void BuilderAddCellTest2()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1");
            builder.AddCell("A1", "C1", ExternalCellValue.DataType.Formula);
            builder.AddCell("B2", "TRUE", ExternalCellValue.DataType.Boolean);
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets[0].Cells.Count);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("C1", result.Worksheets[0].Cells[new Address("A1")].Value);
            Assert.Equal(ExternalCellValue.DataType.Formula, result.Worksheets[0].Cells[new Address("A1")].Type);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("B2")]);
            Assert.Equal("TRUE", result.Worksheets[0].Cells[new Address("B2")].Value);
            Assert.Equal(ExternalCellValue.DataType.Boolean, result.Worksheets[0].Cells[new Address("B2")].Type);
        }

        [Fact(DisplayName = "Test of the failing external link builder function for adding cells on no worksheets")]
        public void BuilderAddCellFailTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.Throws<ArgumentException>(() => { builder.AddCell("A1", "test"); });
        }

        [Fact(DisplayName = "Test of the failing external link builder function overload for adding cells on no worksheets")]
        public void BuilderAddCellFailTest2()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.Throws<ArgumentException>(() => { builder.AddCell("A1", "42", ExternalCellValue.DataType.Number); });
        }


        [Theory(DisplayName = "Test of the failing external link builder function for adding cells on invalid values")]
        [InlineData("ZZZZZ1", "test")]
        [InlineData("A0", "test")]
        [InlineData("A99999999999999", "test")]
        [InlineData("A-5", "test")]
        [InlineData("_A1", "test")]
        [InlineData(null, "test")]
        [InlineData("", "test")]
        [InlineData("0", "test")]
        public void BuilderAddCellFailTest3(string address, string value)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder.AddWorksheet("ws1");
            Assert.ThrowsAny<Exception>(() => { builder.AddCell(address, value); });
        }


        #endregion


    }
}

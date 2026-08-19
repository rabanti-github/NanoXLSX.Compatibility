using NanoXLSX.Extensions;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Reader;
using System.IO;
using System.IO.Packaging;
using Xunit;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Compatibility.Test.Reader
{
    public class ExternalLinkReaderTest
    {
        [Fact(DisplayName = "Test of the unused properties for null (for coverage)")]
        public void NoOpPropertiesTest()
        {
            ExternalLinkReader reader = new ExternalLinkReader();
            reader.Init(new MemoryStream(), new Workbook(), new ReaderOptions(), null);
            Assert.Null(reader.StreamEntryName);
        }

        [Fact(DisplayName = "Test reading external-link URI roles, worksheets, and refresh states")]
        public void ReadsUriRolesWorksheetsAndRefreshStates()
        {
            RelationshipCatalog catalog = ExternalLinkTestUtils.CreateRelationshipCatalog(target: @"..\data\external.xlsx");
            ExternalLinkTestUtils.AddExternalRelationship(catalog, "rId2", @"C:\data\external.xlsx");
            ExternalLinkTestUtils.AddExternalRelationship(catalog, "rId3", @"alternate\external.xlsx");
            const string alternateUrls =
                "<xxl21:alternateUrls><xxl21:absoluteUrl r:id=\"rId2\"/>" +
                "<xxl21:relativeUrl r:id=\"rId3\"/></xxl21:alternateUrls>";
            const string sheetNames =
                "<sheetNames><sheetName val=\"First\"/><sheetName val=\"Second\"/></sheetNames>";
            const string sheetData =
                "<sheetDataSet><sheetData sheetId=\"0\" refreshErrors=\"1\"/>" +
                "<sheetData sheetId=\"1\" refreshErrors=\"0\"/></sheetDataSet>";
            string xml = ExternalLinkTestUtils.CreateExternalLinkXml(
                alternateUrls: alternateUrls,
                sheetNames: sheetNames,
                sheetDataSet: sheetData);

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkReader(xml, catalog);

            ExternalLink link = Assert.Single(workbook.GetExternalLinks());
            Assert.Equal(@"..\data\external.xlsx", link.TargetUri);
            Assert.Equal(@"C:\data\external.xlsx", link.AbsoluteAlternateUri);
            Assert.Equal(@"alternate\external.xlsx", link.RelativeAlternateUri);
            Assert.Equal("rIdWorkbookExternal1", link.WorkbookRId);
            Assert.Collection(
                link.Worksheets,
                sheet =>
                {
                    Assert.Equal("First", sheet.Name);
                    Assert.True(sheet.RefreshErros);
                },
                sheet =>
                {
                    Assert.Equal("Second", sheet.Name);
                    Assert.False(sheet.RefreshErros);
                });
        }

        [Theory(DisplayName = "Test reading cached external-cell data types")]
        [InlineData(null, "42", ExternalCellValue.DataType.Number)]
        [InlineData("b", "1", ExternalCellValue.DataType.Boolean)]
        [InlineData("d", "2026-08-19", ExternalCellValue.DataType.Date)]
        [InlineData("e", "#N/A", ExternalCellValue.DataType.Error)]
        [InlineData("s", "shared text", ExternalCellValue.DataType.String)]
        [InlineData("str", "text", ExternalCellValue.DataType.String)]
        [InlineData("unknown", "42", ExternalCellValue.DataType.Number)]
        public void ReadsCachedExternalCellType(string xmlType, string value, ExternalCellValue.DataType expectedType)
        {
            string xml = ExternalLinkTestUtils.CreateCachedCellXml(xmlType, value);

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkReader(
                xml,
                ExternalLinkTestUtils.CreateRelationshipCatalog());

            ExternalCellValue cell = Assert.Single(Assert.Single(workbook.GetExternalLinks()).Worksheets).Cells[new Address("A1")];
            Assert.Equal(expectedType, cell.Type);
            Assert.Equal(value, cell.Value);
        }

        [Fact(DisplayName = "Test reading cached cell metadata and an empty cached cell")]
        public void ReadsCellMetadataAndEmptyCell()
        {
            string xml = ExternalLinkTestUtils.CreateCachedCellXml(null, null, " vm=\"7\"");

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkReader(
                xml,
                ExternalLinkTestUtils.CreateRelationshipCatalog());

            ExternalCellValue cell = Assert.Single(Assert.Single(workbook.GetExternalLinks()).Worksheets).Cells[new Address("A1")];
            Assert.Equal(ExternalCellValue.DataType.Empty, cell.Type);
            Assert.Equal(string.Empty, cell.Value);
            Assert.Equal(7, cell.CellMetadata);
        }

        [Fact(DisplayName = "Test reading cached external defined names")]
        public void ReadsCachedExternalDefinedNames()
        {
            string xml = ExternalLinkTestUtils.CreateExternalLinkXml(
                definedNames:
                    "<definedNames><definedName name=\"First\" refersTo=\"Data!$A$1\" sheetId=\"2\"/>" +
                    "<definedName name=\"WithoutReference\"/></definedNames>");

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkReader(
                xml,
                ExternalLinkTestUtils.CreateRelationshipCatalog());

            Assert.Collection(
                Assert.Single(workbook.GetExternalLinks()).DefinedNames,
                name =>
                {
                    Assert.Equal("First", name.Name);
                    Assert.Equal("Data!$A$1", name.RefersTo);
                    Assert.Equal("2", name.RelationshipId);
                },
                name =>
                {
                    Assert.Equal("WithoutReference", name.Name);
                    Assert.Null(name.RefersTo);
                    Assert.Null(name.RelationshipId);
                });
        }

        [Fact(DisplayName = "Test reading an external link without optional URI relationships")]
        public void ReadsLinkWithoutOptionalRelationships()
        {
            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkReader(
                ExternalLinkTestUtils.CreateExternalLinkXml(),
                ExternalLinkTestUtils.CreateRelationshipCatalog());

            ExternalLink link = Assert.Single(workbook.GetExternalLinks());
            Assert.Null(link.AbsoluteAlternateUri);
            Assert.Null(link.RelativeAlternateUri);
        }

        [Fact(DisplayName = "Test failure when the external-link relationship catalog is missing")]
        public void RejectsMissingRelationshipCatalog()
        {
            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    ExternalLinkTestUtils.CreateExternalLinkXml(),
                    null));

            Assert.Contains("relationship catalog", exception.InnerException.Message);
        }

        [Fact(DisplayName = "Test failure when the required external-link target relationship ID is missing")]
        public void RejectsMissingTargetRelationshipId()
        {
            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    ExternalLinkTestUtils.CreateExternalLinkXml(targetRelationshipId: null),
                    ExternalLinkTestUtils.CreateRelationshipCatalog()));

            Assert.Contains("relationship ID is missing", exception.InnerException.Message);
        }

        [Fact(DisplayName = "Test failure when an external-link relationship cannot be resolved")]
        public void RejectsUnresolvedRelationshipId()
        {
            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    ExternalLinkTestUtils.CreateExternalLinkXml(targetRelationshipId: "rIdMissing"),
                    ExternalLinkTestUtils.CreateRelationshipCatalog()));

            Assert.Contains("could not be resolved", exception.InnerException.Message);
        }

        [Theory(DisplayName = "Test failure for an invalid external-link relationship definition")]
        [InlineData("invalid-type", TargetMode.External)]
        [InlineData(ExternalLinkTestUtils.ExternalLinkPathRelationshipType, TargetMode.Internal)]
        public void RejectsInvalidRelationship(string relationshipType, TargetMode targetMode)
        {
            RelationshipCatalog catalog = ExternalLinkTestUtils.CreateRelationshipCatalog(
                relationshipType: relationshipType,
                targetMode: targetMode);

            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    ExternalLinkTestUtils.CreateExternalLinkXml(),
                    catalog));

            Assert.Contains("invalid type or target mode", exception.InnerException.Message);
        }

        [Theory(DisplayName = "Test failure for invalid cached external worksheet XML")]
        [InlineData("<sheetNames/>", null)]
        [InlineData("<sheetNames><sheetName val=\"Data\"/></sheetNames>", "<sheetDataSet><sheetData sheetId=\"9\"/></sheetDataSet>")]
        [InlineData("<sheetNames><sheetName val=\"Data\"/></sheetNames>", "<sheetDataSet><sheetData sheetId=\"0\"><row><cell r=\"INVALID\"><v>1</v></cell></row></sheetData></sheetDataSet>")]
        public void RejectsInvalidCachedWorksheetXml(string sheetNames, string sheetDataSet)
        {
            string xml = ExternalLinkTestUtils.CreateExternalLinkXml(
                sheetNames: sheetNames,
                sheetDataSet: sheetDataSet);

            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    xml,
                    ExternalLinkTestUtils.CreateRelationshipCatalog()));

            Assert.NotNull(exception.InnerException);
        }

        [Fact(DisplayName = "Test wrapping malformed external-link XML")]
        public void WrapsMalformedExternalLinkXml()
        {
            IOException exception = Assert.Throws<IOException>(() =>
                ExternalLinkTestUtils.ExecuteExternalLinkReader(
                    "<externalLink><externalBook>",
                    ExternalLinkTestUtils.CreateRelationshipCatalog()));

            Assert.NotNull(exception.InnerException);
            Assert.Contains("stream", exception.Message);
        }
    }
}

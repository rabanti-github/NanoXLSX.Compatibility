using NanoXLSX.Exceptions;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Reader;
using NanoXLSX.Registry;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Reader
{
    public class ExternalLinkWorkbookReaderTest
    {
        [Fact(DisplayName = "Test reading workbook external-reference relationship IDs in order")]
        public void ReadsExternalReferenceRelationshipIdsInOrder()
        {
            const string xml =
                "<workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<externalReferences><externalReference r:id=\"rId4\"/><externalReference/>" +
                "<externalReference r:id=\"rId9\"/></externalReferences></workbook>";

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkWorkbookReader(xml);

            List<string> ids = workbook.AuxiliaryData.GetData<List<string>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY);
            Assert.Equal(new[] { "rId4", "rId9" }, ids);
        }

        [Fact(DisplayName = "Test reading global and local external defined-name references")]
        public void ReadsGlobalAndLocalExternalDefinedNames()
        {
            const string xml =
                "<workbook><definedNames>" +
                "<definedName name=\"GlobalName\"><![CDATA[[1]Data!$A$1]]></definedName>" +
                "<definedName name=\"LocalName\" localSheetId=\"2\" xml:space=\"preserve\"> [2]Other!$B$2 </definedName>" +
                "<definedName name=\"Empty\"/>" +
                "<definedName name=\"Readable\">C:\\data\\[book.xlsx]Data!A1</definedName>" +
                "</definedNames></workbook>";

            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkWorkbookReader(xml);

            List<ExternalDefinedNameReference> names = workbook.AuxiliaryData.GetData<List<ExternalDefinedNameReference>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY);
            Assert.Collection(
                names,
                name =>
                {
                    Assert.Equal("GlobalName", name.Name);
                    Assert.Null(name.LocalSheetIndex);
                    Assert.Equal("[1]Data!$A$1", name.Expression);
                },
                name =>
                {
                    Assert.Equal("LocalName", name.Name);
                    Assert.Equal(2, name.LocalSheetIndex);
                    Assert.Equal(" [2]Other!$B$2 ", name.Expression);
                });
        }

        [Fact(DisplayName = "Test that absent workbook external-link sections create no auxiliary data")]
        public void CreatesNoDataForAbsentExternalLinkSections()
        {
            Workbook workbook = ExternalLinkTestUtils.ExecuteExternalLinkWorkbookReader("<workbook><sheets/></workbook>");

            Assert.Null(workbook.AuxiliaryData.GetData<List<string>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY));
            Assert.Null(workbook.AuxiliaryData.GetData<List<ExternalDefinedNameReference>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY));
        }

        [Theory(DisplayName = "Test wrapping invalid workbook external-link XML")]
        [InlineData("<workbook><externalReferences>")]
        [InlineData("<workbook><definedNames><definedName name=\"Bad\" localSheetId=\"x\">[1]Data!A1</definedName></definedNames></workbook>")]
        public void WrapsInvalidWorkbookExternalLinkXml(string xml)
        {
            IOException exception = Assert.Throws<IOException>(
                () => ExternalLinkTestUtils.ExecuteExternalLinkWorkbookReader(xml));

            Assert.NotNull(exception.InnerException);
            Assert.Contains("stream", exception.Message);
        }
    }
}

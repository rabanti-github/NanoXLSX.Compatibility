using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Registry;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class CompatibilityPreparingInlineWriteProcessorTest
    {
        [Fact(DisplayName = "Test of the resolution of multiple links for defined names")]
        public void ResolvesMultipleLinksForDefinedNamesTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)";
            DefinedName definedName = workbook.AddDefinedNameFormula("ExternalTotal", expression);

            Execute(workbook, "C:\\temp\\book one.xlsx", "../other.xlsx");

            const string expected = "SUM('[1]Sheet 1'!$A$1,'[2]Data'!$B$2)";

            Assert.Equal(expected, GetResolvedDefinedName(workbook, 0).Expression);
            Assert.Equal(new[] { 1, 2 }, GetResolvedDefinedName(workbook, 0).LinkIndexes);
            Assert.Equal(expression, definedName.TextValue);
        }

        [Fact(DisplayName = "Test of the resolution of multiple links for cell formulas")]
        public void ResolvesMultipleLinksForCellFormulasTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);

            Execute(workbook, "C:\\temp\\book one.xlsx", "../other.xlsx");

            const string expected = "SUM('[1]Sheet 1'!$A$1,'[2]Data'!$B$2)";

            Assert.Equal(expected, GetResolvedFormula(workbook, 0, "A1").Expression);
            Assert.Equal(new[] { 1, 2 }, GetResolvedFormula(workbook, 0, "A1").LinkIndexes);
            Assert.Equal(expression, cell.Value);
            Assert.Equal(expression, cell.Formula.Expression);
        }

        [Theory]
        [DisplayName("Test of path variants for defined names")]
        [InlineData("book.xlsx", "[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:\\temp\\book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:/temp/book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("../data/book.xlsx", "..\\data\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("/mnt/data/book.xlsx", "/mnt/data/[book.xlsx]Sheet1!$A$1")]
        [InlineData("file:///C:/temp/book%20one.xlsx", "C:\\temp\\[book one.xlsx]Sheet1!$A$1")]
        [InlineData("file://server/share/book.xlsx", "\\\\server\\share\\[book.xlsx]Sheet1!$A$1")]
        public void ResolvesPathVariantsDefinedNamesTest(string linkUri, string formula)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddDefinedNameFormula("definedName", formula);
            Execute(workbook, linkUri);

            Assert.Equal("[1]Sheet1!$A$1", GetResolvedDefinedName(workbook, 0).Expression);
        }

        [Theory]
        [DisplayName("Test of path variants for cell formulas")]
        [InlineData("book.xlsx", "[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:\\temp\\book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:/temp/book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("../data/book.xlsx", "..\\data\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("/mnt/data/book.xlsx", "/mnt/data/[book.xlsx]Sheet1!$A$1")]
        [InlineData("file:///C:/temp/book%20one.xlsx", "C:\\temp\\[book one.xlsx]Sheet1!$A$1")]
        [InlineData("file://server/share/book.xlsx", "\\\\server\\share\\[book.xlsx]Sheet1!$A$1")]
        public void ResolvesPathVariantsCellFormulasTest(string linkUri, string formula)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(formula, "A1");
            Execute(workbook, linkUri);

            Assert.Equal("[1]Sheet1!$A$1", GetResolvedFormula(workbook, 0, "A1").Expression);
        }

        [Fact(DisplayName = "Test of the resolution with distinct cases of links")]
        public void ResolvesCaseDistinctLinksByExactCaseTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "/data/[Book.xlsx]Upper!A1+/data/[book.xlsx]Lower!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");

            Execute(workbook,
                "/data/Book.xlsx",
                "/data/book.xlsx");

            const string expected = "[1]Upper!A1+[2]Lower!A1";
            var resolved = GetResolvedFormula(workbook, 0, "A1");
            Assert.Equal(expected, resolved.Expression);
            Assert.Equal(2, resolved.LinkIndexes.Count);
        }

        [Fact(DisplayName = "Test of summary of multiple identical links")]
        public void ResolvesCaseInsensitiveLinksTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "/data/[Book.xlsx]Upper!A1+/data/[Book.xlsx]Upper2!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");

            Execute(workbook,
                "/data/book.xlsx",
                "/data/Book.xlsx");

            const string expected = "[2]Upper!A1+[2]Upper2!A1"; // 2nd index from added links
            var resolved = GetResolvedFormula(workbook, 0, "A1");
            Assert.Equal(expected, resolved.Expression);
            Assert.Single(resolved.LinkIndexes);
        }

        [Fact(DisplayName = "Test of the failed attempt to resolve a link when multiple links could match")]
        public void RejectsNonExactCaseWhenMultipleLinksCouldMatchTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("/data/[BOOK.xlsx]Sheet1!A1", "A1");

            NotSupportedContentException exception = Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                "/data/Book.xlsx",
                "/data/book.xlsx"));

            Assert.Contains("Sheet1!A1", exception.Message);
            Assert.Contains("/data/[BOOK.xlsx]", exception.Message);
            Assert.Contains("1, 2", exception.Message);
        }

        [Fact(DisplayName = "Test of the failed resolution if multiple identical links where defined")]
        public void RejectsIdenticalTargetsTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddDefinedNameFormula("ExternalName", "/data/[book.xlsx]Sheet1!A1");

            NotSupportedContentException exception = Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                "/data/book.xlsx",
                "/data/book.xlsx"));

            Assert.Contains("defined name 'ExternalName'", exception.Message);
            Assert.Contains("1, 2", exception.Message);
        }

        [Fact(DisplayName = "Test of the resolution of files with identical names but different directories")]
        public void DistinguishesSameFilenameInDifferentDirectoriesTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(
                "/first/[book.xlsx]A!A1+/second/[book.xlsx]B!B2",
                "A1");

            Execute(workbook,
                "/first/book.xlsx",
                "/second/book.xlsx");

            Assert.Equal(
                "[1]A!A1+[2]B!B2",
                GetResolvedFormula(workbook, 0, "A1").Expression);
        }

        [Fact(DisplayName = "Test of the resolution of external links when multiple URIs for the same links are defined")]
        public void CoalescesMultipleAliasesOfOneLinkTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            ExternalLink link = new ExternalLink();
            link.SetReadUris("../book.xlsx", null, "..\\book.xlsx");

            Execute(workbook, link);

            Assert.Equal("[1]Sheet1!A1", GetResolvedFormula(workbook, 0, "A1").Expression);
        }

        [Fact(DisplayName = "Test of formula resolution against all URI roles of one link")]
        public void ResolvesAllUriRolesOfOneLinkTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(
                "C:\\Files\\[book.xlsx]One!A1+..\\[book.xlsx]Two!A1+alternative\\[book.xlsx]Three!A1",
                "A1");
            ExternalLink link = new ExternalLink();
            link.SetReadUris(@"..\book.xlsx", @"C:\Files\book.xlsx", @"alternative\book.xlsx");

            Execute(workbook, link);

            Assert.Equal(
                "[1]One!A1+[1]Two!A1+[1]Three!A1",
                GetResolvedFormula(workbook, 0, "A1").Expression);
        }

        [Fact(DisplayName = "Test of the failed resolution when multiple URIs for a link are defined but the only difference are ambiguous path separators")]
        public void RejectsSeparatorAliasesThatCollapseAcrossLinksTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");

            Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                "../book.xlsx",
                "..\\book.xlsx"));
        }

        [Fact(DisplayName = "Test of non-resolution of links on cell formulas if the formula expression was set to a defined name afterwards")]
        public void UsesDefinedNameReferenceInsteadOfFormulaExpressionTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName definedName = workbook.AddDefinedNameFormula("LocalFormula", "1+1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.DefinedNameReference = definedName; // Forces the formula to use a defined name

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolvedFormula(workbook, 0, "A1"));
            Assert.Null(GetResolvedDefinedName(workbook, 0)); // Also not expected
            Assert.Same(definedName, cell.Formula.DefinedNameReference);
        }

        [Fact(DisplayName = "Test of the resolution of external links on multiple worksheets for defined names")]
        public void UsesCollisionFreeKeysForScopedNamesTest()
        {
            Workbook workbook = new Workbook();
            Worksheet firstSheet = new Worksheet("First");
            Worksheet secondSheet = new Worksheet("Second");
            workbook.AddWorksheet(firstSheet);
            workbook.AddWorksheet(secondSheet);
            const string expression = "..\\[book.xlsx]Data!A1";
            workbook.AddDefinedNameFormula("ScopedName", expression, firstSheet);
            workbook.AddDefinedNameFormula("ScopedName", expression, secondSheet);

            Execute(workbook, "../book.xlsx");

            Assert.Equal("[1]Data!A1", GetResolvedDefinedName(workbook, 0).Expression);
            Assert.Equal("[1]Data!A1", GetResolvedDefinedName(workbook, 1).Expression);
        }

        [Fact(DisplayName = "Test of the resolution of external links on multiple worksheets for cell formulas")]
        public void UsesCollisionFreeKeysForWorksheetCellsTest()
        {
            Workbook workbook = new Workbook();
            Worksheet firstSheet = new Worksheet("First");
            Worksheet secondSheet = new Worksheet("Second");
            workbook.AddWorksheet(firstSheet);
            workbook.AddWorksheet(secondSheet);
            const string expression = "..\\[book.xlsx]Data!A1";
            firstSheet.AddCellFormula(expression, "A1");
            secondSheet.AddCellFormula(expression, "A1");

            Execute(workbook, "../book.xlsx");

            Assert.Equal("[1]Data!A1", GetResolvedFormula(workbook, 0, "A1").Expression);
            Assert.Equal("[1]Data!A1", GetResolvedFormula(workbook, 1, "A1").Expression);
        }


        [Fact(DisplayName = "Test of non-resolution of external links on cell formulas when the formula was (forced) to null")]
        public void DoesNotFallBackToCellValueWhenFormulaExpressionIsNullTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string value = "..\\[book.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(value, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.Expression = null;

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolvedFormula(workbook, 0, "A1"));
            Assert.Null(GetResolvedDefinedName(workbook, 0)); // Also not expected
            Assert.Equal(value, cell.Value);
            Assert.Null(cell.Formula.Expression);
        }

        [Fact(DisplayName = "Test of the non-resolution of formulas without links and unrelated cell content (strings)")]
        public void IgnoresUnmatchedEmptyAndNonFormulaContent()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("..\\[book.xlsx]Sheet1!A1", "A1");
            workbook.CurrentWorksheet.AddCellFormula("SUM(A1:A2)", "A2");

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolvedFormula(workbook, 0, "A1"));
            Assert.Null(GetResolvedFormula(workbook, 0, "A2"));

            Workbook noLinksWorkbook = new Workbook("Sheet1");
            noLinksWorkbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Execute(noLinksWorkbook, new string[0]);
            Assert.Null(GetResolvedFormula(noLinksWorkbook, 0, "A1"));
            Assert.Null(GetResolvedDefinedName(noLinksWorkbook, 0)); // Also not expected
        }

        [Theory(DisplayName = "Test of resolution handling of corner case URIs as external links")]
        [InlineData(@"file://server/share/directory/file.xlsx")]
        [InlineData(@"/C:/directory/file.xlsx")]
        [InlineData(@"file:///C%3A/directory/file.xlsx")]
        [InlineData(@"https://example.com/path/file.xlsx")]
        [InlineData(@"relative/directory/[file.xlsx]")]
        [InlineData(@"https://example.com/path/[file.xlsx]")]
        public void CornerCaseUrisTest(string uri)
        {
            Workbook workbook = new Workbook("Sheet1");
            Execute(workbook, uri);

            Assert.Null(GetResolvedFormula(workbook, 0, "A1"));
            Assert.Null(GetResolvedDefinedName(workbook, 0));
            Assert.Null(GetResolvedDefinedName(workbook, 1));
        }

        #region helperMethods

        private static void Execute(Workbook workbook, params string[] links)
        {
            ExternalLink[] externalLinks = new ExternalLink[links.Length];
            for (int i = 0; i < links.Length; i++)
            {
                ExternalLink link = new ExternalLink();
                link.SetReadUris(links[i], null, null);
                externalLinks[i] = link;
            }
            Execute(workbook, externalLinks);
        }

        private static void Execute(Workbook workbook, params ExternalLink[] externalLinks)
        {
            for (int i = 0; i < externalLinks.Length; i++)
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                    i,
                    externalLinks[i],
                    true
                    );
            }
            CompatibilityPreparingInlineWriteProcessor processor = new CompatibilityPreparingInlineWriteProcessor();
            processor.Init(new TestWriteContext(workbook));
            processor.Execute();
        }

        private static ExternalLinkResolution GetResolvedFormula(Workbook workbook, int worksheetIndex, string cellAddress)
        {
            Dictionary<int, Dictionary<string, ExternalLinkResolution>> resolutions =
                workbook.AuxiliaryData.GetData<Dictionary<int, Dictionary<string, ExternalLinkResolution>>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY);
            if (resolutions == null || !resolutions.TryGetValue(worksheetIndex, out Dictionary<string, ExternalLinkResolution> worksheetResolutions))
            {
                return null;
            }
            worksheetResolutions.TryGetValue(cellAddress, out ExternalLinkResolution resolution);
            return resolution;
        }

        private static ExternalLinkResolution GetResolvedDefinedName(Workbook workbook, int definedNameIndex)
        {
            Dictionary<int, ExternalLinkResolution> resolutions =
                workbook.AuxiliaryData.GetData<Dictionary<int, ExternalLinkResolution>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY);
            if (resolutions == null)
            {
                return null;
            }
            resolutions.TryGetValue(definedNameIndex, out ExternalLinkResolution resolution);
            return resolution;
        }

        private sealed class TestWriteContext : IWriteContext
        {
            public Workbook Workbook { get; }
            [ExcludeFromCodeCoverage]
            public IWriterProcessingData WriterProcessingData => null;

            public TestWriteContext(Workbook workbook)
            {
                Workbook = workbook;
            }

            [ExcludeFromCodeCoverage]
            public void MarkFeatureAsPrepared(string featureUuid)
            {
            }

            [ExcludeFromCodeCoverage]
            public bool IsFeaturePrepared(string featureUuid)
            {
                return false;
            }
        }

        #endregion
    }
}

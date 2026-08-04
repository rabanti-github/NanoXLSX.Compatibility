using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal;
using NanoXLSX.Registry;
using System.Collections.Generic;
using System.ComponentModel;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class CompatibilityPreparingInlineWriteProcessorTest
    {
        [Fact (DisplayName ="Test of the resolution of multiple links for defined names")]
        public void ResolvesMultipleLinksForDefinedNamesTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)";
            DefinedName definedName = workbook.AddDefinedNameFormula("ExternalTotal", expression);

            Execute(workbook, "C:\\temp\\book one.xlsx", "../other.xlsx");

            const string expected = "SUM('[1]Sheet 1'!$A$1,'[2]Data'!$B$2)";

            string defNameEntityId = CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY;

            Assert.Equal(expected, GetResolved(workbook, defNameEntityId, "0").Expression);
            Assert.Equal(expected, GetResolved(workbook, defNameEntityId, "0").Expression);
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

            string formulaEntityId = CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY;

            Assert.Equal(expected, GetResolved(workbook, formulaEntityId, "0:A1").Expression);
            Assert.Equal(expected, GetResolved(workbook, formulaEntityId, "0:A1").Expression);
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

            Assert.Equal("[1]Sheet1!$A$1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "0").Expression);
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

            Assert.Equal("[1]Sheet1!$A$1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
        }

        [Fact (DisplayName ="Test of the resolution with distinct cases of links")]
        public void ResolvesCaseDistinctLinksByExactCaseTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "/data/[Book.xlsx]Upper!A1+/data/[book.xlsx]Lower!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");

            Execute(workbook,
                "/data/Book.xlsx",
                "/data/book.xlsx");

            const string expected = "[1]Upper!A1+[2]Lower!A1";
            var resolved = GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1");
            Assert.Equal(expected, resolved.Expression);
            Assert.Equal(2, resolved.LinkIndexes.Count);
        }

        [Fact (DisplayName ="Test of summary of multiple identical links")]
        public void ResolvesCaseInsensitiveLinksTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "/data/[Book.xlsx]Upper!A1+/data/[Book.xlsx]Upper2!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");

            Execute(workbook,
                "/data/book.xlsx",
                "/data/Book.xlsx");

            const string expected = "[2]Upper!A1+[2]Upper2!A1"; // 2nd index from added links
            var resolved = GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1");
            Assert.Equal(expected, resolved.Expression);
            Assert.Equal(1, resolved.LinkIndexes.Count);
        }

        [Fact (DisplayName ="Test of the failed attempt to resolve a link when multiple links could match")]
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

        [Fact (DisplayName ="Test of the resolution of files with identical names but different directories")]
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
                GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
        }

        [Fact (DisplayName ="Test of the resolution of external links when multiple URIs for the same links are defined")]
        public void CoalescesMultipleAliasesOfOneLinkTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            List<string> uris = new List<string> { "../book.xlsx", "..\\book.xlsx" }; // Multiple URIs in one ext. link

            Execute(workbook, uris);

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
        }

        [Fact (DisplayName ="Test of the failed resolution when multiple URIs for a link are defined but the only difference are ambiguous path separators")]
        public void RejectsSeparatorAliasesThatCollapseAcrossLinksTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");

            Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                "../book.xlsx",
                "..\\book.xlsx"));
        }

        [Fact (DisplayName ="Test of the preference of a FormulaData object on the resolution of links on cell formulas")]
        public void PrefersNestedFormulaExpressionWhenCellValueDiffersTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string formulaExpression = "C:\\formula\\[book.xlsx]Sheet1!A1";
            const string cellValue = "C:\\value\\[other.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(formulaExpression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Value = cellValue;

            Execute(workbook, "C:\\formula\\book.xlsx");

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
            Assert.Equal(cellValue, cell.Value);
            Assert.Equal(formulaExpression, cell.Formula.Expression);
        }

        [Fact (DisplayName ="Test of the fallback to the cell value when no FormulaData object was defined, on the resolution of links on cell formulas")]
        public void FallsBackToCellValueOnlyWhenFormulaObjectIsNullTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string expression = "..\\[book.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula = null;

            Execute(workbook, "../book.xlsx");

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
            Assert.Equal(expression, cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact (DisplayName ="Test of non-resolution of links on cell formulas if the formula expression was set to a defined name afterwards")]
        public void UsesDefinedNameReferenceInsteadOfFormulaExpressionTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName definedName = workbook.AddDefinedNameFormula("LocalFormula", "1+1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.DefinedNameReference = definedName; // Forces the formula to use a defined name

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1"));
            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "0")); // Also not expected
            Assert.Same(definedName, cell.Formula.DefinedNameReference);
        }

        [Fact (DisplayName ="Test of the resolution of external links on multiple worksheets for defined names")]
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

            Assert.Equal("[1]Data!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "0").Expression);
            Assert.Equal("[1]Data!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "1").Expression);
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

            Assert.Equal("[1]Data!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1").Expression);
            Assert.Equal("[1]Data!A1", GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "1:A1").Expression);
        }


        [Fact (DisplayName ="Test of non-resolution of external links on cell formulas when the formula was (forced) to null")]
        public void DoesNotFallBackToCellValueWhenFormulaExpressionIsNullTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string value = "..\\[book.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(value, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.Expression = null;

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1"));
            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "0")); // Also not expected
            Assert.Equal(value, cell.Value);
            Assert.Null(cell.Formula.Expression);
        }

        [Fact (DisplayName ="Test of the non-resolution of formulas without links and unrelated cell content (strings)")]
        public void IgnoresUnmatchedEmptyAndNonFormulaContent()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("..\\[book.xlsx]Sheet1!A1", "A1");
            workbook.CurrentWorksheet.AddCellFormula("SUM(A1:A2)", "A2");

            Execute(workbook, "../book.xlsx");

            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1"));
            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A2"));

            Workbook noLinksWorkbook = new Workbook("Sheet1");
            noLinksWorkbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Execute(noLinksWorkbook, new string[0]);
            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY, "0:A1"));
            Assert.Null(GetResolved(workbook, CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY, "0")); // Also not expected
        }

        #region helperMethods

        private static void Execute(Workbook workbook, params string[] links)
        {
            List<string>[] uriLists = new List<string>[links.Length];
            for (int i = 0; i < links.Length; i++)
            {
                uriLists[i] = new List<string> { links[i] };
            }
            Execute(workbook, uriLists);
        }

        private static void Execute(Workbook workbook, params List<string>[] uriLists)
        {
            int i = 0;
            foreach (List<string> uriList in uriLists)
            {
                ExternalLink link = new ExternalLink();
                ExternalLinkBuilder builder = new ExternalLinkBuilder(link);
                foreach (string uri in uriList)
                {
                    builder.AddUri(uri);
                }
                link = builder.Build();
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                    i,
                    link,
                    true
                    );
                i++;
            }
            CompatibilityPreparingInlineWriteProcessor processor = new CompatibilityPreparingInlineWriteProcessor();
            processor.Init(new TestWriteContext(workbook));
            processor.Execute();
        }

        private static ExternalLinkResolution GetResolved(Workbook workbook, string entityId, string valueId)
        {
            return workbook.AuxiliaryData.GetData<ExternalLinkResolution>(
                PlugInUUID.CompatibilityInlineProcessor,
                entityId,
                valueId);
        }

        private sealed class TestWriteContext : IWriteContext
        {
            public Workbook Workbook { get; }
            public IWriterProcessingData WriterProcessingData => null;

            public TestWriteContext(Workbook workbook)
            {
                Workbook = workbook;
            }

            public void MarkFeatureAsPrepared(string featureUuid)
            {
            }

            public bool IsFeaturePrepared(string featureUuid)
            {
                return false;
            }
        }

        #endregion
    }
}

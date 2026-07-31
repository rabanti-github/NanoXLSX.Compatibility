using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal;
using NanoXLSX.Registry;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class CompatibilityPreparingInlineWriteProcessorTest
    {
        [Fact]
        public void ResolvesMultipleLinksForDefinedNamesAndCellFormulas()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)";
            DefinedName definedName = workbook.AddDefinedNameFormula("ExternalTotal", expression);
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);

            Execute(workbook,
                new ExternalLink("C:\\temp\\book one.xlsx"),
                new ExternalLink("../other.xlsx"));

            const string expected = "SUM('[1]Sheet 1'!$A$1,'[2]Data'!$B$2)";
            Assert.Equal(expected, GetResolved(workbook, "1", "definedName:0"));
            Assert.Equal(expected, GetResolved(workbook, "2", "definedName:0"));
            Assert.Equal(expected, GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal(expected, GetResolved(workbook, "2", "cell:0:A1"));
            Assert.Equal(expression, definedName.TextValue);
            Assert.Equal(expression, cell.Value);
            Assert.Equal(expression, cell.Formula.Expression);
        }

        [Theory]
        [InlineData("book.xlsx", "[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:\\temp\\book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("C:/temp/book.xlsx", "C:\\temp\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("../data/book.xlsx", "..\\data\\[book.xlsx]Sheet1!$A$1")]
        [InlineData("/mnt/data/book.xlsx", "/mnt/data/[book.xlsx]Sheet1!$A$1")]
        [InlineData("file:///C:/temp/book%20one.xlsx", "C:\\temp\\[book one.xlsx]Sheet1!$A$1")]
        [InlineData("file://server/share/book.xlsx", "\\\\server\\share\\[book.xlsx]Sheet1!$A$1")]
        public void ResolvesPathVariants(string linkUri, string formula)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(formula, "A1");

            Execute(workbook, new ExternalLink(linkUri));

            Assert.Equal("[1]Sheet1!$A$1", GetResolved(workbook, "1", "cell:0:A1"));
        }

        [Fact]
        public void ResolvesCaseDistinctLinksByExactCase()
        {
            Workbook workbook = new Workbook("Sheet1");
            string expression = "/data/[Book.xlsx]Upper!A1+/data/[book.xlsx]Lower!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");

            Execute(workbook,
                new ExternalLink("/data/Book.xlsx"),
                new ExternalLink("/data/book.xlsx"));

            const string expected = "[1]Upper!A1+[2]Lower!A1";
            Assert.Equal(expected, GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal(expected, GetResolved(workbook, "2", "cell:0:A1"));
        }

        [Fact]
        public void RejectsNonExactCaseWhenMultipleLinksCouldMatch()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("/data/[BOOK.xlsx]Sheet1!A1", "A1");

            NotSupportedContentException exception = Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                new ExternalLink("/data/Book.xlsx"),
                new ExternalLink("/data/book.xlsx")));

            Assert.Contains("Sheet1!A1", exception.Message);
            Assert.Contains("/data/[BOOK.xlsx]", exception.Message);
            Assert.Contains("1, 2", exception.Message);
        }

        [Fact]
        public void RejectsIdenticalTargets()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddDefinedNameFormula("ExternalName", "/data/[book.xlsx]Sheet1!A1");

            NotSupportedContentException exception = Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                new ExternalLink("/data/book.xlsx"),
                new ExternalLink("/data/book.xlsx")));

            Assert.Contains("defined name 'ExternalName'", exception.Message);
            Assert.Contains("1, 2", exception.Message);
        }

        [Fact]
        public void DistinguishesSameFilenameInDifferentDirectories()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(
                "/first/[book.xlsx]A!A1+/second/[book.xlsx]B!B2",
                "A1");

            Execute(workbook,
                new ExternalLink("/first/book.xlsx"),
                new ExternalLink("/second/book.xlsx"));

            Assert.Equal(
                "[1]A!A1+[2]B!B2",
                GetResolved(workbook, "1", "cell:0:A1"));
        }

        [Fact]
        public void CoalescesMultipleAliasesOfOneLink()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            ExternalLink link = new ExternalLink("../book.xlsx");
            link.CreateBuilder().AddUri("..\\book.xlsx");

            Execute(workbook, link);

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, "1", "cell:0:A1"));
        }

        [Fact]
        public void RejectsSeparatorAliasesThatCollapseAcrossLinks()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");

            Assert.Throws<NotSupportedContentException>(() => Execute(
                workbook,
                new ExternalLink("../book.xlsx"),
                new ExternalLink("..\\book.xlsx")));
        }

        [Fact]
        public void PrefersNestedFormulaExpressionWhenCellValueDiffers()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string formulaExpression = "C:\\formula\\[book.xlsx]Sheet1!A1";
            const string cellValue = "C:\\value\\[other.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(formulaExpression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Value = cellValue;

            Execute(workbook, new ExternalLink("C:\\formula\\book.xlsx"));

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal(cellValue, cell.Value);
            Assert.Equal(formulaExpression, cell.Formula.Expression);
        }

        [Fact]
        public void FallsBackToCellValueOnlyWhenFormulaObjectIsNull()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string expression = "..\\[book.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula = null;

            Execute(workbook, new ExternalLink("../book.xlsx"));

            Assert.Equal("[1]Sheet1!A1", GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal(expression, cell.Value);
            Assert.Null(cell.Formula);
        }

        [Fact]
        public void UsesDefinedNameReferenceInsteadOfFormulaExpression()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName definedName = workbook.AddDefinedNameFormula("LocalFormula", "1+1");
            workbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.DefinedNameReference = definedName;

            Execute(workbook, new ExternalLink("../book.xlsx"));

            Assert.Null(GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Same(definedName, cell.Formula.DefinedNameReference);
        }

        [Fact]
        public void UsesCollisionFreeKeysForScopedNamesAndWorksheetCells()
        {
            Workbook workbook = new Workbook();
            Worksheet firstSheet = new Worksheet("First");
            Worksheet secondSheet = new Worksheet("Second");
            workbook.AddWorksheet(firstSheet);
            workbook.AddWorksheet(secondSheet);
            const string expression = "..\\[book.xlsx]Data!A1";
            workbook.AddDefinedNameFormula("ScopedName", expression, firstSheet);
            workbook.AddDefinedNameFormula("ScopedName", expression, secondSheet);
            firstSheet.AddCellFormula(expression, "A1");
            secondSheet.AddCellFormula(expression, "A1");

            Execute(workbook, new ExternalLink("../book.xlsx"));

            Assert.Equal("[1]Data!A1", GetResolved(workbook, "1", "definedName:0"));
            Assert.Equal("[1]Data!A1", GetResolved(workbook, "1", "definedName:1"));
            Assert.Equal("[1]Data!A1", GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal("[1]Data!A1", GetResolved(workbook, "1", "cell:1:A1"));
        }

        [Fact]
        public void DoesNotFallBackToCellValueWhenFormulaExpressionIsNull()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string value = "..\\[book.xlsx]Sheet1!A1";
            workbook.CurrentWorksheet.AddCellFormula(value, "A1");
            Cell cell = workbook.CurrentWorksheet.GetCell(0, 0);
            cell.Formula.Expression = null;

            Execute(workbook, new ExternalLink("../book.xlsx"));

            Assert.Null(GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Equal(value, cell.Value);
            Assert.Null(cell.Formula.Expression);
        }

        [Fact]
        public void IgnoresUnmatchedEmptyAndNonFormulaContent()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCell("..\\[book.xlsx]Sheet1!A1", "A1");
            workbook.CurrentWorksheet.AddCellFormula("SUM(A1:A2)", "A2");

            Execute(workbook, new ExternalLink("../book.xlsx"));

            Assert.Null(GetResolved(workbook, "1", "cell:0:A1"));
            Assert.Null(GetResolved(workbook, "1", "cell:0:A2"));

            Workbook noLinksWorkbook = new Workbook("Sheet1");
            noLinksWorkbook.CurrentWorksheet.AddCellFormula("..\\[book.xlsx]Sheet1!A1", "A1");
            Execute(noLinksWorkbook);
            Assert.Null(GetResolved(noLinksWorkbook, "1", "cell:0:A1"));
        }

        private static void Execute(Workbook workbook, params ExternalLink[] links)
        {
            for (int i = 0; i < links.Length; i++)
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    "a",
                    i,
                    links[i],
                    true);
            }
            CompatibilityPreparingInlineWriteProcessor processor = new CompatibilityPreparingInlineWriteProcessor();
            processor.Init(new TestWriteContext(workbook));
            processor.Execute();
        }

        private static string GetResolved(Workbook workbook, string entityId, string valueId)
        {
            return workbook.AuxiliaryData.GetData<string>(
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
    }
}

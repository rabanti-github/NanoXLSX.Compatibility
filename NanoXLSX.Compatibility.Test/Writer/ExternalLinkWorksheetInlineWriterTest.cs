using NanoXLSX.Extensions;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class ExternalLinkWorksheetInlineWriterTest
    {
        [Fact(DisplayName = "Test of the unused properties for null (for coverage)")]
        public void NoOpPropertiesTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("SUM(A1:A2)", "A1");
            XmlElement root = CreateWorksheetRoot("A1", "SUM(A1:A2)");
            ExternalLinkWorksheetInlineWriter writer = CreateWriter(workbook, ref root, workbook.CurrentWorksheet.SheetID);

            Assert.Null(writer.XmlElement);
            Assert.Null(writer.WriteContext);
        }

        [Fact(DisplayName = "Test that worksheet XML remains unchanged without external links")]
        public void LeavesWorksheetXmlUnchangedWithoutExternalLinks()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula("SUM(A1:A2)", "A1");
            XmlElement root = CreateWorksheetRoot("A1", "SUM(A1:A2)");
            ExternalLinkWorksheetInlineWriter writer = CreateWriter(workbook, ref root, workbook.CurrentWorksheet.SheetID);

            writer.Execute();

            Assert.Equal("SUM(A1:A2)", GetFormula(root, "A1").InnerValue);
        }

        [Fact(DisplayName = "Test replacement of a resolved external-link worksheet formula")]
        public void ReplacesResolvedWorksheetFormulaWithoutMutatingWorkbook()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string original = @"C:\data\[external.xlsx]Data!A1";
            workbook.CurrentWorksheet.AddCellFormula(original, "A1");
            StoreResolutions(workbook, workbook.CurrentWorksheet.SheetID, "A1", "[1]Data!A1");
            XmlElement root = CreateWorksheetRoot("A1", original);
            ExternalLinkWorksheetInlineWriter writer = CreateWriter(workbook, ref root, workbook.CurrentWorksheet.SheetID);

            writer.Execute();

            Assert.Equal("[1]Data!A1", GetFormula(root, "A1").InnerValue);
            Assert.Equal(original, workbook.CurrentWorksheet.GetCell(0, 0).Formula.Expression);
        }

        // TODO split up in three tests
        [Theory(DisplayName = "Test ignoring unavailable worksheet formula resolutions")]
        [InlineData("missing")]
        [InlineData("empty")]
        [InlineData("other-sheet")]
        public void IgnoresUnavailableFormulaResolutions(string scenario)
        {
            Workbook workbook = new Workbook("Sheet1");
            const string original = @"C:\data\[external.xlsx]Data!A1";
            workbook.CurrentWorksheet.AddCellFormula(original, "A1");
            if (scenario == "empty")
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY,
                    new Dictionary<int, Dictionary<string, ExternalLinkResolution>>());
            }
            else if (scenario == "other-sheet")
            {
                StoreResolutions(workbook, workbook.CurrentWorksheet.SheetID + 50, "A1", "[1]Data!A1");
            }
            XmlElement root = CreateWorksheetRoot("A1", original);
            ExternalLinkWorksheetInlineWriter writer = CreateWriter(workbook, ref root, workbook.CurrentWorksheet.SheetID);

            writer.Execute();

            Assert.Equal(original, GetFormula(root, "A1").InnerValue);
        }

        [Theory(DisplayName = "Test ignoring a resolution whose worksheet XML node is unavailable")]
        [InlineData(false)]
        [InlineData(true)]
        public void IgnoresUnavailableWorksheetXmlNode(bool includeCellWithoutFormula)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.CurrentWorksheet.AddCellFormula(@"C:\data\[external.xlsx]Data!A1", "A1");
            StoreResolutions(workbook, workbook.CurrentWorksheet.SheetID, "A1", "[1]Data!A1");
            XmlElement root = XmlElement.CreateElement("worksheet");
            if (includeCellWithoutFormula)
            {
                XmlElement cell = XmlElement.CreateElement("c");
                cell.AddAttribute("r", "A1");
                root.AddChildElement(cell);
            }
            ExternalLinkWorksheetInlineWriter writer = CreateWriter(workbook, ref root, workbook.CurrentWorksheet.SheetID);

            writer.Execute();

            Assert.Empty(root.FindChildElementsByName("f"));
        }

        [Fact(DisplayName = "Test failure when initializing the worksheet writer with an unknown sheet ID")]
        public void RejectsUnknownWorksheetId()
        {
            Workbook workbook = new Workbook("Sheet1");
            XmlElement root = XmlElement.CreateElement("worksheet");
            ExternalLinkWorksheetInlineWriter writer = new ExternalLinkWorksheetInlineWriter();

            Assert.Throws<InvalidOperationException>(() => writer.Init(ref root, workbook, 999));
        }

        private static ExternalLinkWorksheetInlineWriter CreateWriter(Workbook workbook, ref XmlElement root, int sheetId)
        {
            ExternalLinkWorksheetInlineWriter writer = new ExternalLinkWorksheetInlineWriter();
            writer.Init(ref root, workbook, sheetId);
            return writer;
        }

        private static XmlElement CreateWorksheetRoot(string address, string expression)
        {
            XmlElement root = XmlElement.CreateElement("worksheet");
            XmlElement cell = XmlElement.CreateElement("c");
            cell.AddAttribute("r", address);
            XmlElement formula = XmlElement.CreateElement("f");
            formula.InnerValue = expression;
            cell.AddChildElement(formula);
            root.AddChildElement(cell);
            return root;
        }

        private static XmlElement GetFormula(XmlElement root, string address)
        {
            XmlElement cell = Assert.Single(root.FindChildElementsByNameAndAttribute("c", "r", address));
            return Assert.Single(cell.FindChildElementsByName("f"));
        }

        private static void StoreResolutions(Workbook workbook, int sheetId, string address, string expression)
        {
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY,
                new Dictionary<int, Dictionary<string, ExternalLinkResolution>>
                {
                    [sheetId] = new Dictionary<string, ExternalLinkResolution>
                    {
                        [address] = new ExternalLinkResolution(expression, new List<int> { 1 })
                    }
                });
        }
    }
}

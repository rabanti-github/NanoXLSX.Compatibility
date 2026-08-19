using NanoXLSX.Extensions;
using NanoXLSX.Registry;
using System.Linq;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Reader
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class ExternalLinkRoundTripTest
    {
        [Fact(DisplayName = "Test round trip of external links, formulas, defined names, and cached data")]
        public void RoundTripsExternalLinkWorkbookDataWithoutMutatingSource()
        {
            PlugInLoader.DisposePlugins();
            PlugInLoader.Initialize();
            Workbook source = new Workbook("Sheet1");
            const string cellExpression = @"C:\data\[first.xlsx]Data!A1+D:\other\[second.xlsx]Other!B2";
            const string definedNameExpression = @"C:\data\[first.xlsx]Data!$A$1";
            const string loadedCellExpression =
                "[file:///C:/data/first.xlsx]Data!A1+[file:///D:/other/second.xlsx]Other!B2";
            const string loadedDefinedNameExpression = "[file:///C:/data/first.xlsx]Data!$A$1";
            source.CurrentWorksheet.AddCellFormula(cellExpression, "A1");
            DefinedName sourceDefinedName = source.AddDefinedNameFormula("ExternalTotal", definedNameExpression);

            ExternalLink first = new ExternalLink(@"C:\data\first.xlsx", @"..\data\first.xlsx");
            ExternalWorksheet cachedSheet = new ExternalWorksheet("Data");
            cachedSheet.RefreshErros = true;
            cachedSheet.AddCell("A1", "cached", ExternalCellValue.DataType.String);
            cachedSheet.Cells[new Address("A1")].CellMetadata = 5;
            cachedSheet.AddCell("A2", "1", ExternalCellValue.DataType.Boolean);
            cachedSheet.AddCell("A3", null, ExternalCellValue.DataType.Empty);
            first.AddWorksheet(cachedSheet);
            ExternalDefinedName externalDefinedName = new ExternalDefinedName("CachedName", "Data!$A$1");
            externalDefinedName.RelationshipId = "0";
            first.AddDefinedName(externalDefinedName);
            ExternalLink second = new ExternalLink(@"D:\other\second.xlsx");
            source.AddExternalLink(first);
            source.AddExternalLink(second);

            Workbook loaded = ExternalLinkTestUtils.RoundTrip(source);

            Assert.Equal(cellExpression, source.CurrentWorksheet.GetCell(0, 0).Formula.Expression);
            Assert.Equal(definedNameExpression, sourceDefinedName.TextValue);
            Assert.Equal(loadedCellExpression, loaded.CurrentWorksheet.GetCell(0, 0).Formula.Expression);
            Assert.Equal(
                loadedDefinedNameExpression,
                loaded.GetDefinedNames().Single(name => name.Name == "ExternalTotal").TextValue);
            Assert.Collection(
                loaded.GetExternalLinks(),
                link =>
                {
                    Assert.Equal("../data/first.xlsx", link.TargetUri);
                    Assert.Equal("file:///C:/data/first.xlsx", link.AbsoluteAlternateUri);
                    ExternalWorksheet sheet = Assert.Single(link.Worksheets);
                    Assert.True(sheet.RefreshErros);
                    Assert.Equal("cached", sheet.Cells[new Address("A1")].Value);
                    Assert.Equal(5, sheet.Cells[new Address("A1")].CellMetadata);
                    Assert.Equal(ExternalCellValue.DataType.Boolean, sheet.Cells[new Address("A2")].Type);
                    Assert.Equal(ExternalCellValue.DataType.Empty, sheet.Cells[new Address("A3")].Type);
                    ExternalDefinedName name = Assert.Single(link.DefinedNames);
                    Assert.Equal("CachedName", name.Name);
                    Assert.Equal("Data!$A$1", name.RefersTo);
                    Assert.Equal("0", name.RelationshipId);
                },
                link => Assert.Equal("file:///D:/other/second.xlsx", link.TargetUri));
        }
    }
}

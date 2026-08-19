using System;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalWorksheetTest
    {

        [Fact(DisplayName = "Test of the properties and the constructor")]
        public void PropertyConstructorTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            Assert.Equal("name", ws.Name);
            Assert.NotNull(ws.Cells);
            Assert.Empty(ws.Cells);
        }

        [Theory(DisplayName = "Test of the failing constructor on invalid values")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("\\")]
        [InlineData("[test]")]
        [InlineData("xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx")]
        [InlineData("???")]
        [InlineData("*")]
        public void ConstructorFailTest(string name)
        {
            Assert.ThrowsAny<Exception>(() => { ExternalWorksheet ws = new ExternalWorksheet(name); });
        }

        [Fact(DisplayName = "Test of the AddCell method")]
        public void AddCellTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A2", "test1");
            ws.AddCell("A1", "test2");
            Assert.NotEmpty(ws.Cells);
            Assert.Equal(2, ws.Cells.Count);
            Assert.Equal("test1", ws.Cells[new Address("A2")].Value);
            Assert.Equal("test2", ws.Cells[new Address("A1")].Value);
            Assert.Equal(ExternalCellValue.DataType.String, ws.Cells[new Address("A2")].Type); // Default
            Assert.Equal(ExternalCellValue.DataType.String, ws.Cells[new Address("A1")].Type); // "
        }

        [Fact(DisplayName = "Test of the AddCell method when overwriting existing cells")]
        public void AddCellOverwriteTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test1", ExternalCellValue.DataType.Date);
            ws.AddCell("A1", "newValue", ExternalCellValue.DataType.String);
            Assert.NotEmpty(ws.Cells);
            Assert.Single(ws.Cells);
            Assert.Equal("newValue", ws.Cells[new Address("A1")].Value);
            Assert.Equal(ExternalCellValue.DataType.String, ws.Cells[new Address("A1")].Type);
        }

        [Fact(DisplayName = "Test of the AddCell method (overload)")]
        public void AddCellTest2()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A2", "0", ExternalCellValue.DataType.Boolean);
            ws.AddCell("A1", "55", ExternalCellValue.DataType.Number);
            Assert.NotEmpty(ws.Cells);
            Assert.Equal(2, ws.Cells.Count);
            Assert.Equal("0", ws.Cells[new Address("A2")].Value);
            Assert.Equal("55", ws.Cells[new Address("A1")].Value);
            Assert.Equal(ExternalCellValue.DataType.Boolean, ws.Cells[new Address("A2")].Type);
            Assert.Equal(ExternalCellValue.DataType.Number, ws.Cells[new Address("A1")].Type);
        }

        [Theory(DisplayName = "Test of the failing AddCell method on invalid values (address)")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("\t")]
        [InlineData("ZZZZZ1")]
        [InlineData("A0")]
        [InlineData("A99999999999999")]
        [InlineData("A-5")]
        [InlineData("_A1")]
        public void AddCellFailTest(string address)
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            Assert.ThrowsAny<Exception>(() => { ws.AddCell(address, "test"); });
        }

        [Theory(DisplayName = "Test of the failing AddCell (overload) method on invalid values (address)")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("\t")]
        [InlineData("ZZZZZ1")]
        [InlineData("A0")]
        [InlineData("A99999999999999")]
        [InlineData("A-5")]
        [InlineData("_A1")]
        public void AddCellFailTest2(string address)
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            Assert.ThrowsAny<Exception>(() => { ws.AddCell(address, "test", ExternalCellValue.DataType.Boolean); });
        }


        [Fact(DisplayName = "Test of the RemoveCell method")]
        public void RemoveCellTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test1");
            ws.AddCell("A2", "test2");
            Assert.NotEmpty(ws.Cells);
            Assert.Equal(2, ws.Cells.Count);
            bool check = ws.RemoveCell("A1");
            Assert.Single(ws.Cells);
            Assert.True(check);
            Assert.Equal("test2", ws.Cells[new Address("A2")].Value);

            check = ws.RemoveCell("A2");
            Assert.True(check);
            Assert.Empty(ws.Cells);
        }

        [Fact(DisplayName = "Test of the RemoveCell method on unknown cells or a empty list")]
        public void RemoveCellUnknownEmptyTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test1");
            Assert.NotEmpty(ws.Cells);
            Assert.Single(ws.Cells);
            bool check = ws.RemoveCell("A2");
            Assert.Single(ws.Cells);
            Assert.False(check);
            Assert.Equal("test1", ws.Cells[new Address("A1")].Value); // Should be remaining

            ws.RemoveCell("A1");
            Assert.Empty(ws.Cells);

            check = ws.RemoveCell("A1");
            Assert.False(check);
            Assert.Empty(ws.Cells);

            check = ws.RemoveCell("A99");
            Assert.False(check);
            Assert.Empty(ws.Cells);
        }

        [Theory(DisplayName = "Test of the failing RemoveCell method on invalid values (address)")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("\t")]
        [InlineData("ZZZZZ1")]
        [InlineData("A0")]
        [InlineData("A99999999999999")]
        [InlineData("A-5")]
        [InlineData("_A1")]
        public void RemoveCellFailTest(string address)
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test");
            Assert.ThrowsAny<Exception>(() => { ws.RemoveCell(address); });
        }

        [Fact(DisplayName = "Test of the TryGet method")]
        public void TryGetTest()
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test1");
            ws.AddCell("A2", "test2");

            ExternalCellValue cell;
            bool check = ws.TryGetCell("A1", out cell);
            Assert.NotNull(cell);
            Assert.True(check);
            Assert.Equal("test1", cell.Value);

            ExternalCellValue cell2;
            check = ws.TryGetCell("A3", out cell2);
            Assert.False(check);
            Assert.Null(cell2);
        }

        [Theory(DisplayName = "Test of the TryGetCell method on invalid values (address); should not throw")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("\t")]
        [InlineData("ZZZZZ1")]
        [InlineData("A0")]
        [InlineData("A99999999999999")]
        [InlineData("A-5")]
        [InlineData("_A1")]
        public void TryGetCellTest2(string address)
        {
            ExternalWorksheet ws = new ExternalWorksheet("name");
            ws.AddCell("A1", "test");
            ExternalCellValue cell;

            bool check = ws.TryGetCell(address, out cell);
            Assert.False(check);
            Assert.Null(cell);
        }

    }
}

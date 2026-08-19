using System;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalCellValueTest
    {

        [Theory(DisplayName = "Test of the property handling")]
        [InlineData(null, "", ExternalCellValue.DataType.Empty)]
        [InlineData("test", "test", ExternalCellValue.DataType.String)]
        public void PropertiesTest(string value, string expectedValue, ExternalCellValue.DataType expectedType)
        {
            ExternalCellValue cellValue = new ExternalCellValue(value);
            Assert.Equal(expectedValue, cellValue.Value);
            Assert.Equal(expectedType, cellValue.Type);
            Assert.Null(cellValue.CellMetadata);

            cellValue.CellMetadata = 22;
            Assert.Equal(22, cellValue.CellMetadata);
        }

        [Theory(DisplayName = "Test of the constructor handling")]
        [InlineData("0", "0", ExternalCellValue.DataType.Number)]
        [InlineData("99", "99", ExternalCellValue.DataType.Number)]
        [InlineData("", "", ExternalCellValue.DataType.Empty)]
        [InlineData(null, "", ExternalCellValue.DataType.Empty)]
        [InlineData("", "", ExternalCellValue.DataType.Error)]
        [InlineData("test", "test", ExternalCellValue.DataType.String)]
        [InlineData("1", "1", ExternalCellValue.DataType.Boolean)]
        [InlineData("true", "1", ExternalCellValue.DataType.Boolean)]
        [InlineData("0", "0", ExternalCellValue.DataType.Boolean)]
        [InlineData("false", "0", ExternalCellValue.DataType.Boolean)]
        [InlineData("1587", "1587", ExternalCellValue.DataType.Date)]
        [InlineData("A5", "A5", ExternalCellValue.DataType.String)] // Formula = string
        public void ConstructorTest(string given, string expected, ExternalCellValue.DataType type)
        {
            ExternalCellValue cellValue = new ExternalCellValue(given, type);
            Assert.Equal(expected, cellValue.Value);
            Assert.Equal(type, cellValue.Type);
        }

        [Theory(DisplayName = "Test of the failing constructor handling")]
        [InlineData("test", ExternalCellValue.DataType.Boolean)]
        [InlineData("", ExternalCellValue.DataType.Boolean)]
        [InlineData("2", ExternalCellValue.DataType.Boolean)]
        public void FailingConstructorTest(string given, ExternalCellValue.DataType type)
        {
            Assert.ThrowsAny<Exception>(() => new ExternalCellValue(given, type));

        }

    }
}

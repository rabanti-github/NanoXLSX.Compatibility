using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalCellValueTest
    {

        [Fact(DisplayName = "Test of the property handling")]
        public void PropertiesTest()
        {
            ExternalCellValue cellValue = new ExternalCellValue(null);
            Assert.Equal("", cellValue.Value); // Default behavior
            Assert.Equal(ExternalCellValue.DataType.Empty, cellValue.Type); // Default behavior

            ExternalCellValue cellValue2 = new ExternalCellValue("test");
            Assert.Equal("test", cellValue2.Value);
            Assert.Equal(ExternalCellValue.DataType.SharedString, cellValue2.Type); // Default behavior

        }

        [Theory(DisplayName = "Test of the constructor handling")]
        [InlineData("0", "0", ExternalCellValue.DataType.Number)]
        [InlineData("99", "99", ExternalCellValue.DataType.Number)]
        [InlineData("", "", ExternalCellValue.DataType.Empty)]
        [InlineData(null, "", ExternalCellValue.DataType.Empty)]
        [InlineData("", "", ExternalCellValue.DataType.Error)]
        [InlineData("test", "test", ExternalCellValue.DataType.SharedString)]
        [InlineData("inline", "inline", ExternalCellValue.DataType.InlineString)]
        [InlineData("TRUE", "TRUE", ExternalCellValue.DataType.Boolean)]
        [InlineData("1587", "1587", ExternalCellValue.DataType.Date)]
        [InlineData("A5", "A5", ExternalCellValue.DataType.Formula)]
        public void ConstructorTest(string given, string expected, ExternalCellValue.DataType type)
        {
            ExternalCellValue cellValue = new ExternalCellValue(given, type);
            Assert.Equal(expected, cellValue.Value);
            Assert.Equal(type, cellValue.Type);
        }

    }
}

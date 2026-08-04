using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalDefinedNameTest
    {
        [Fact(DisplayName = "Test of the property handling")]
        public void PropertiesTest()
        {
            ExternalDefinedName name = new ExternalDefinedName("a", "b");
            Assert.Equal("a", name.Name);
            Assert.Equal("b", name.RefersTo);
        }

        [Theory(DisplayName = "Test of the constructor handling")]
        [InlineData("0", "0")]
        [InlineData("99", "99")]
        [InlineData("name", "C5")]
        [InlineData("link", "workbook.xlsx")]
        public void ConstructorTest(string name, string refersTo)
        {
            ExternalDefinedName definedName = new ExternalDefinedName(name, refersTo);
            Assert.Equal(name, definedName.Name);
            Assert.Equal(refersTo, definedName.RefersTo);
        }

        [Theory(DisplayName = "Test of the failing constructor handling on invalid values")]
        [InlineData("", "")]
        [InlineData(null, "")]
        [InlineData("", null)]
        [InlineData(" ", "a")]
        [InlineData("a", " ")]
        [InlineData("a", null)]
        [InlineData(null, "a")]
        [InlineData(" ", null)]
        [InlineData(null, " ")]
        [InlineData("a", "\t")]
        [InlineData("   ", "a")]
        public void FailingConstructorTest(string name, string refersTo)
        {
            Assert.Throws<FormatException>(() => { ExternalDefinedName definedName = new ExternalDefinedName(name, refersTo); });
        }

    }
}

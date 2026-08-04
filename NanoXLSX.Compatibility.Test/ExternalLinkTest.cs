using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalLinkTest
    {
        [Fact(DisplayName = "Test of the properties")]
        public void PropertiesTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("extName");

            link.AddUri("uri");
            link.AddDefinedName("name", "link");
            link.AddWorksheet(ws);

            Assert.Single(link.Uris);
            Assert.Single(link.DefinedNames);
            Assert.Single(link.Worksheets);

            Assert.Equal("uri", link.Uris[0]);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("extName", link.Worksheets[0].Name);
        }

        [Fact(DisplayName = "Test of the default constructor")]
        public void ConstructorTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.NotNull(link.Uris);
            Assert.NotNull(link.DefinedNames);
            Assert.NotNull(link.Worksheets);
            Assert.Empty(link.Uris);
            Assert.Empty(link.DefinedNames);
            Assert.Empty(link.Worksheets);
        }

        [Fact(DisplayName = "Test of the named constructor")]
        public void ConstructorTest2()
        {
            ExternalLink link = new ExternalLink("uriName");
            Assert.Single(link.Uris);
            Assert.Empty(link.DefinedNames);
            Assert.Empty(link.Worksheets);
            Assert.Equal("uriName", link.Uris[0]);
        }

        [Theory(DisplayName = "Test of the failing named constructor on invalid values")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t")]
        public void FailingConstructorTest(string uri)
        {
            Assert.Throws<ArgumentException>(() => { ExternalLink link = new ExternalLink(uri); });
        }

        [Theory(DisplayName = "Test of the AddUri method")]
        [InlineData("a")]
        [InlineData(@"file://server")]
        [InlineData(@"file://server/share/directory/file.xlsx")]
        [InlineData(@"//server")]
        [InlineData(@"//server/")]
        [InlineData(@"/C:/directory/file.xlsx")]
        [InlineData(@"file:///C%3A/directory/file.xlsx")]
        [InlineData(@"https://example.com/path/file.xlsx")]
        [InlineData(@"relative/directory/")]
        [InlineData(@"file:///C:/directory/")]
        [InlineData(@"file://server/share/directory/")]
        [InlineData(@"relative/directory/[file.xlsx]")]
        [InlineData(@"https://example.com/path/[file.xlsx]")]
        public void AddUriTest(string uri)
        {
            ExternalLink link = new ExternalLink();
            link.AddUri(uri);
            Assert.Single(link.Uris);
            Assert.Equal(uri, link.Uris[0]); 
        }

        [Fact(DisplayName = "Test of the AddUri method on multiple values")]
        public void AddMultipleUrisTest()
        {
            ExternalLink link = new ExternalLink();
            link.AddUri("a");
            link.AddUri("b");
            Assert.Equal(2, link.Uris.Count);
            Assert.Equal("a", link.Uris[0]);
            Assert.Equal("b", link.Uris[1]);
        }

        [Theory(DisplayName = "Test of the failing AddUri method on invalid values")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t")]
        public void FailingAddUriTest(string uri)
        {
            ExternalLink link = new ExternalLink();
            Assert.Throws<ArgumentException>(() => { link.AddUri(uri); });
        }

        [Fact(DisplayName = "Test of the AddDefinedName method")]
        public void AddDefinedNameTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            link.AddDefinedName(name);
            Assert.Single(link.DefinedNames);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("refTo", link.DefinedNames[0].RefersTo);
        }

        [Fact(DisplayName = "Test of the AddDefinedName method on multiple values")]
        public void AddMultipleDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            ExternalDefinedName name2 = new ExternalDefinedName("name2", "refTo2");
            link.AddDefinedName(name);
            link.AddDefinedName(name2);
            Assert.Equal(2, link.DefinedNames.Count);
            Assert.Equal("name", link.DefinedNames[0].Name);
            Assert.Equal("refTo", link.DefinedNames[0].RefersTo);
            Assert.Equal("name2", link.DefinedNames[1].Name);
            Assert.Equal("refTo2", link.DefinedNames[1].RefersTo);
        }

        [Fact(DisplayName = "Test of the failing AddDefinedName method on null")]
        public void FailingAddDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.Throws<ArgumentException>(() => { link.AddDefinedName(null); });
        }

        [Fact(DisplayName = "Test of the failing AddDefinedName method on duplicate values")]
        public void FailingDuplicateDefinedNamesTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalDefinedName name = new ExternalDefinedName("name", "refTo");
            link.AddDefinedName(name);
            ExternalDefinedName name2 = new ExternalDefinedName("name", "refTo2");
            Assert.Throws<ArgumentException>(() => { link.AddDefinedName(name2); });
        }

        [Fact(DisplayName = "Test of the AddWorksheet method")]
        public void AddWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Single(link.Worksheets);
            Assert.Equal("ext", link.Worksheets[0].Name);
        }

        [Fact(DisplayName = "Test of the AddWorksheet method on multiple values")]
        public void AddMultipleWorksheetsTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            ExternalWorksheet ws2 = new ExternalWorksheet("ext2");
            link.AddWorksheet(ws);
            link.AddWorksheet(ws2);
            Assert.Equal(2, link.Worksheets.Count);
            Assert.Equal("ext", link.Worksheets[0].Name);
            Assert.Equal("ext2", link.Worksheets[1].Name);
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet method on null")]
        public void FailingAddWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            Assert.Throws<ArgumentException>(() => { link.AddWorksheet(null); });
        }

        [Fact(DisplayName = "Test of the failing AddWorksheet method on duplicate values")]
        public void FailingDuplicateWorksheetsTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            ExternalWorksheet ws2 = new ExternalWorksheet("ext");
            Assert.Throws<ArgumentException>(() => { link.AddWorksheet(ws2); });
        }

        [Fact(DisplayName = "Test of the GetWorksheet method")]
        public void GetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            ExternalWorksheet ws2 = new ExternalWorksheet("ext2");
            link.AddWorksheet(ws);
            link.AddWorksheet(ws2);
            ExternalWorksheet restult = link.GetWorksheet("ext");
            Assert.NotNull(restult);
            Assert.Equal("ext", restult.Name);
        }

        [Fact(DisplayName = "Test of the failing GetWorksheet method on null")]
        public void FailingGetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Throws<ArgumentException>(() => { link.GetWorksheet(null); });
        }

        [Fact(DisplayName = "Test of the failing GetWorksheet on a not existing name")]
        public void FailingNotExistingGetWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalWorksheet ws = new ExternalWorksheet("ext");
            link.AddWorksheet(ws);
            Assert.Throws<ArgumentException>(() => { link.GetWorksheet("ext2"); });
        }

        [Fact(DisplayName = "Test of the CreateBuilder method")]
        public void CreateBuilderTest()
        {
            ExternalLink link = new ExternalLink();
            var builder = link.CreateBuilder();
            Assert.NotNull(builder);
            Assert.IsType<ExternalLinkBuilder>(builder);
        }


    }
}

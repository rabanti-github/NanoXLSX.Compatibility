using System;
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

        #region builderTests

        [Fact(DisplayName = "Test of the CreateBuilder method in ExternalLink")]
        public void CreateBuilderTest()
        {
            ExternalLink link = new ExternalLink();
            var builder = link.CreateBuilder();
            Assert.NotNull(builder);
            Assert.IsType<ExternalLinkBuilder>(builder);
        }

        [Fact(DisplayName = "Test of the external link passing function of the builder")]
        public void BuilderAddExternalLinkTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            ExternalLink extLink = new ExternalLink("uri1");
            ExternalLinkBuilder passed = new ExternalLinkBuilder(extLink);
            ExternalLink result = passed.Build();
            Assert.NotNull(result);
            Assert.Equal("uri1", result.Uris[0]);
        }

        [Fact(DisplayName = "Test of the failing external link passing function of the builder on null")]
        public void BuilderAddExternalLinkFailTest()
        {
            Assert.Throws<ArgumentException>(() => { var builder = new ExternalLinkBuilder(null); });
        }

        [Fact(DisplayName = "Test of the external link builder function for URIs")]
        public void BuilderUriTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddUri("uri1").AddUri("uri2");
            ExternalLink result = builder.Build();
            Assert.Equal(2, result.Uris.Count);
            Assert.Equal("uri1", result.Uris[0]);
            Assert.Equal("uri2", result.Uris[1]);
        }

        [Theory(DisplayName = "Test of the failing external link builder function for URIs on invalid values")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t")]
        public void BuilderFailingUriTest(string uri)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.Throws<ArgumentException>(() => { builder.AddUri("a").AddUri(uri); });
        }

        [Fact(DisplayName = "Test of the external link builder function for adding worksheets")]
        public void BuilderWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            ExternalLink result = builder.Build();
            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
        }

        [Theory(DisplayName = "Test of the failing external link builder function for adding worksheets on invalid values")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        [InlineData("   ")]
        [InlineData("\t")]
        [InlineData("a")] // duplicate
        public void BuilderFailingWorksheetTest(string name)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.ThrowsAny<Exception>(() => { builder.AddWorksheet("a").AddWorksheet(name); });
        }

        [Fact(DisplayName = "Test of the external link builder function for using worksheets")]
        public void BuilderUseWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            builder.UseWorksheet("ws1");
            builder.AddCell("A1", "test");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
            Assert.Empty(result.Worksheets[1].Cells);
            Assert.Equal(1, result.Worksheets[0].Cells.Count);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[0].Cells[new Address("A1")].Value);
        }

        [Fact(DisplayName = "Test of the external link builder function for using worksheets automatically on add")]
        public void BuilderUsingAddedWorksheetTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            builder.AddCell("A1", "test");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets.Count);
            Assert.Equal("ws1", result.Worksheets[0].Name);
            Assert.Equal("ws2", result.Worksheets[1].Name);
            Assert.Empty(result.Worksheets[0].Cells);
            Assert.Equal(1, result.Worksheets[1].Cells.Count);
            Assert.NotNull(result.Worksheets[1].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[1].Cells[new Address("A1")].Value);
        }

        [Theory(DisplayName = "Test of the failing external link builder function for using worksheets on a invalid value")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData(" ")]
        [InlineData("ws3")]
        public void BuilderUseWorksheetFailTest(string name)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1").AddWorksheet("ws2");
            Assert.ThrowsAny<Exception>(() => { builder.UseWorksheet(name); });
        }

        [Fact(DisplayName = "Test of the external link builder function for adding cells")]
        public void BuilderAddCellTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1");
            builder.AddCell("A1", "test");
            builder.AddCell("B2", "test2");
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets[0].Cells.Count);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("test", result.Worksheets[0].Cells[new Address("A1")].Value);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("B2")]);
            Assert.Equal("test2", result.Worksheets[0].Cells[new Address("B2")].Value);
        }

        [Fact(DisplayName = "Test of the external link builder function for adding cells (overload)")]
        public void BuilderAddCellTest2()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder = builder.AddWorksheet("ws1");
            builder.AddCell("A1", "C1", ExternalCellValue.DataType.Formula);
            builder.AddCell("B2", "TRUE", ExternalCellValue.DataType.Boolean);
            ExternalLink result = builder.Build();

            Assert.Equal(2, result.Worksheets[0].Cells.Count);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("A1")]);
            Assert.Equal("C1", result.Worksheets[0].Cells[new Address("A1")].Value);
            Assert.Equal(ExternalCellValue.DataType.Formula, result.Worksheets[0].Cells[new Address("A1")].Type);
            Assert.NotNull(result.Worksheets[0].Cells[new Address("B2")]);
            Assert.Equal("TRUE", result.Worksheets[0].Cells[new Address("B2")].Value);
            Assert.Equal(ExternalCellValue.DataType.Boolean, result.Worksheets[0].Cells[new Address("B2")].Type);
        }

        [Fact(DisplayName = "Test of the failing external link builder function for adding cells on no worksheets")]
        public void BuilderAddCellFailTest()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.Throws<ArgumentException>(() => { builder.AddCell("A1", "test"); });
        }

        [Fact(DisplayName = "Test of the failing external link builder function overload for adding cells on no worksheets")]
        public void BuilderAddCellFailTest2()
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            Assert.Throws<ArgumentException>(() => { builder.AddCell("A1", "42", ExternalCellValue.DataType.Number); });
        }


        [Theory(DisplayName = "Test of the failing external link builder function for adding cells on invalid values")]
        [InlineData("ZZZZZ1", "test")]
        [InlineData("A0", "test")]
        [InlineData("A99999999999999", "test")]
        [InlineData("A-5", "test")]
        [InlineData("_A1", "test")]
        [InlineData(null, "test")]
        [InlineData("", "test")]
        [InlineData("0", "test")]
        public void BuilderAddCellFailTest3(string address, string value)
        {
            ExternalLink link = new ExternalLink();
            ExternalLinkBuilder builder = link.CreateBuilder();
            builder.AddWorksheet("ws1");
            Assert.ThrowsAny<Exception>(() => { builder.AddCell(address, value); });
        }


        #endregion


    }
}

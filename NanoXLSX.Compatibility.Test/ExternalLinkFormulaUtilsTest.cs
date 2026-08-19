using NanoXLSX.Internal;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalLinkFormulaUtilsTest
    {
        [Theory]
        [InlineData("[1]Sheet1!A1")]
        [InlineData("SUM('[12]Sheet 1'!A1,[2]!ExternalName)")]
        [InlineData("[3]#REF!A1")]
        public void DetectsNumericExternalLinkIds(string expression)
        {
            Assert.True(ExternalLinkFormulaUtils.DetectExternalLinkId(expression));
        }

        [Fact]
        public void DetectsOnlyRequestedNumericExternalLinkId()
        {
            const string expression = "SUM([1]Sheet1!A1,[2]Sheet2!A1)";

            Assert.True(ExternalLinkFormulaUtils.DetectExternalLinkId(expression, "[2]"));
            Assert.False(ExternalLinkFormulaUtils.DetectExternalLinkId(expression, "[3]"));
        }

        [Theory]
        [InlineData("SUM(Table1[1])")]
        [InlineData("\"[1]Sheet1!A1\"")]
        [InlineData("\"Text \"\"[1]Sheet1!A1\"\"\"")]
        [InlineData("[1]")]
        [InlineData("")]
        [InlineData(null)]
        public void IgnoresNonExternalNumericBracketContent(string expression)
        {
            Assert.False(ExternalLinkFormulaUtils.DetectExternalLinkId(expression));
        }

        [Fact]
        public void RejectsInvalidRequestedNumericId()
        {
            Assert.Throws<ArgumentException>(() =>
                ExternalLinkFormulaUtils.DetectExternalLinkId("[1]Sheet1!A1", "[book.xlsx]"));
        }

        [Fact]
        public void ReplacesKnownIdsAndPreservesUnknownIdsAndStrings()
        {
            ExternalLink link = new ExternalLink(@"C:\temp\book.xlsx");
            Dictionary<string, ExternalLink> links = new Dictionary<string, ExternalLink>
            {
                { "[1]", link }
            };
            const string expression = "SUM([1]Sheet1!A1,[2]Sheet2!A1)+\"[1]Text!A1\"";

            string result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(expression, links);

            Assert.Equal("SUM([C:\\temp\\book.xlsx]Sheet1!A1,[2]Sheet2!A1)+\"[1]Text!A1\"", result);
        }

        [Fact]
        public void ReturnsOriginalInstanceWhenNoReplacementIsNeeded()
        {
            string expression = "SUM(A1:A2)";

            string result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(
                expression,
                new Dictionary<string, ExternalLink> { { "[1]", new ExternalLink(@"C:\temp\book.xlsx") } });

            Assert.Same(expression, result);
        }

        [Theory]
        [InlineData(@"C:\temp\[book.xlsx]Sheet1!A1", "[book.xlsx]")]
        [InlineData(@"..\data\[book.xlsx]Data!A1", "[book.xlsx]")]
        [InlineData("https://example.com/data/[book.xlsx]Sheet1!A1", "[book.xlsx]")]
        [InlineData("'[book.xlsx]Sheet 1'!A1", "[book.xlsx]")]
        [InlineData("[book.xlsx]!ExternalName", "[book.xlsx]")]
        [InlineData("SUM([known.xlsx]Sheet1!A1,[unknown.xlsx]Sheet2!A1)", "[known.xlsx]")]
        public void FindsUnresolvedHumanReadableExternalLinks(string expression, string expectedToken)
        {
            Assert.True(ExternalLinkFormulaUtils.TryFindUnresolvedExternalLink(expression, out string token));
            Assert.Equal(expectedToken, token);
        }

        [Theory]
        [InlineData("[1]Sheet1!A1")]
        [InlineData("SUM(Table1[Column])")]
        [InlineData("\"C:\\temp\\[book.xlsx]Sheet1!A1\"")]
        [InlineData("\"Text \"\"[book.xlsx]Sheet1!A1\"\"\"")]
        [InlineData("SUM(A1:A2)")]
        public void IgnoresResolvedAndNonReferenceContent(string expression)
        {
            Assert.False(ExternalLinkFormulaUtils.TryFindUnresolvedExternalLink(expression, out string token));
            Assert.Null(token);
        }
    }
}

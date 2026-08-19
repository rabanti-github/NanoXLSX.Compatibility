using NanoXLSX.Internal;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalLinkFormulaUtilsTest
    {
        [Theory(DisplayName = "Test detection of numeric external link IDs")]
        [InlineData("[1]Sheet1!A1")]
        [InlineData("SUM('[12]Sheet 1'!A1,[2]!ExternalName)")]
        [InlineData("[3]#REF!A1")]
        public void DetectsNumericExternalLinkIds(string expression)
        {
            Assert.True(ExternalLinkFormulaUtils.DetectExternalLinkId(expression));
        }

        [Fact(DisplayName = "Test detection of a requested numeric external link ID")]
        public void DetectsOnlyRequestedNumericExternalLinkId()
        {
            const string expression = "SUM([1]Sheet1!A1,[2]Sheet2!A1)";

            Assert.True(ExternalLinkFormulaUtils.DetectExternalLinkId(expression, "[2]"));
            Assert.False(ExternalLinkFormulaUtils.DetectExternalLinkId(expression, "[3]"));
        }

        [Theory(DisplayName = "Test exclusion of non-external numeric bracket content")]
        [InlineData("SUM(Table1[1])")]
        [InlineData("\"[1]Sheet1!A1\"")]
        [InlineData("\"Text \"\"[1]Sheet1!A1\"\"\"")]
        [InlineData("[1]")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("[")]
        [InlineData("]")]
        public void IgnoresNonExternalNumericBracketContent(string expression)
        {
            Assert.False(ExternalLinkFormulaUtils.DetectExternalLinkId(expression));
        }

        [Fact(DisplayName = "Test rejection of an invalid requested numeric external link ID")]
        public void RejectsInvalidRequestedNumericId()
        {
            Assert.Throws<ArgumentException>(() =>
                ExternalLinkFormulaUtils.DetectExternalLinkId("[1]Sheet1!A1", "[book.xlsx]"));
        }

        [Fact(DisplayName = "Test replacement of known IDs while preserving unknown IDs and string constants")]
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

        [Fact(DisplayName = "Test preservation of IDs after escaped quotes in formula string constants")]
        public void PreservesIdsAfterEscapedQuotesInStringConstants()
        {
            ExternalLink link = new ExternalLink(@"C:\temp\book.xlsx");
            Dictionary<string, ExternalLink> links = new Dictionary<string, ExternalLink>
            {
                { "[1]", link }
            };
            const string expression = "[1]Sheet1!A1+\"Text \"\"quoted\"\" [1]Sheet1!A1\"";

            string result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(expression, links);

            Assert.Equal("[C:\\temp\\book.xlsx]Sheet1!A1+\"Text \"\"quoted\"\" [1]Sheet1!A1\"", result);
        }

        [Fact(DisplayName = "Test rejection of a null external link used for ID replacement")]
        public void RejectsNullExternalLinkReplacement()
        {
            Dictionary<string, ExternalLink> links = new Dictionary<string, ExternalLink>
            {
                { "[1]", null }
            };

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                ExternalLinkFormulaUtils.ReplaceExternalLinkId("[1]Sheet1!A1", links));

            Assert.Equal("links", exception.ParamName);
        }

        [Fact(DisplayName = "Test rejection of an external link without a usable URI for ID replacement")]
        public void RejectsExternalLinkReplacementWithoutUsableUri()
        {
            Dictionary<string, ExternalLink> links = new Dictionary<string, ExternalLink>
            {
                { "[1]", new ExternalLink() }
            };

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                ExternalLinkFormulaUtils.ReplaceExternalLinkId("[1]Sheet1!A1", links));

            Assert.Equal("links", exception.ParamName);
        }

        [Fact(DisplayName = "Test preservation of the original expression when no replacement is needed")]
        public void ReturnsOriginalInstanceWhenNoReplacementIsNeeded()
        {
            string expression = "SUM(A1:A2)";

            string result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(
                expression,
                new Dictionary<string, ExternalLink> { { "[1]", new ExternalLink(@"C:\temp\book.xlsx") } });

            Assert.Same(expression, result);
        }

        [Fact(DisplayName = "Test preservation of the original expression on empty content passed")]
        public void ReturnsOriginalInstanceWhenEmpty()
        {
            string exp = "SUM(A1:A2)";

            // null exp
            string result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(null,
                new Dictionary<string, ExternalLink> { { "[1]", new ExternalLink(@"C:\temp\book.xlsx") } });
            Assert.Null(result);

            // empty exp
            result = ExternalLinkFormulaUtils.ReplaceExternalLinkId("",
                new Dictionary<string, ExternalLink> { { "[1]", new ExternalLink(@"C:\temp\book.xlsx") } });
            Assert.Equal("", result);

            // null links
            result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(exp, null);
            Assert.Same(exp, result);

            // empty links
            result = ExternalLinkFormulaUtils.ReplaceExternalLinkId(exp, new Dictionary<string, ExternalLink>());
            Assert.Same(exp, result);

        }


        [Theory(DisplayName = "Test detection of unresolved human-readable external links")]
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

        [Theory(DisplayName = "Test exclusion of resolved IDs and non-reference content")]
        [InlineData("[1]Sheet1!A1")]
        [InlineData("SUM(Table1[Column])")]
        [InlineData("\"C:\\temp\\[book.xlsx]Sheet1!A1\"")]
        [InlineData("\"Text \"\"[book.xlsx]Sheet1!A1\"\"\"")]
        [InlineData("SUM(A1:A2)")]
        [InlineData("[")]
        [InlineData("]")]
        [InlineData("[Column]+Sheet1!A1")]
        [InlineData("[Column]Sheet1")]
        public void IgnoresResolvedAndNonReferenceContent(string expression)
        {
            Assert.False(ExternalLinkFormulaUtils.TryFindUnresolvedExternalLink(expression, out string token));
            Assert.Null(token);
        }
    }
}

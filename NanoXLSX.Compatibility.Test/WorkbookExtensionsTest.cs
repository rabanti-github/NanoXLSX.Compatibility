using NanoXLSX.Extensions;
using NanoXLSX.Internal;
using NanoXLSX.Registry;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class WorkbookExtensionsTest
    {
        [Fact(DisplayName = "Test of adding an external link")]
        public void AddExternalLinkTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink externalLink = new ExternalLink(@"C:\Files\external.xlsx");

            workbook.AddExternalLink(externalLink);

            ExternalLink result = Assert.Single(workbook.GetExternalLinks());
            Assert.Same(externalLink, result);
        }

        [Fact(DisplayName = "Test of adding an external link when the next index is occupied")]
        public void AddExternalLinkWithOccupiedIndexTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink existing = new ExternalLink(@"C:\Files\existing.xlsx");
            ExternalLink added = new ExternalLink(@"C:\Files\added.xlsx");
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                1,
                existing,
                true);

            workbook.AddExternalLink(added);

            Assert.Same(existing, workbook.AuxiliaryData.GetData<ExternalLink>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                1));
            Assert.Same(added, workbook.AuxiliaryData.GetData<ExternalLink>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                2));
            Assert.Equal(2, workbook.GetExternalLinks().Count);
        }

        [Fact(DisplayName = "Test of adding a null external link")]
        public void AddNullExternalLinkTest()
        {
            Workbook workbook = new Workbook();

            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(
                () => workbook.AddExternalLink((ExternalLink)null));

            Assert.Equal("externalLink", exception.ParamName);
            Assert.Empty(workbook.GetExternalLinks());
        }

        [Fact(DisplayName = "Test of adding an external link using a builder")]
        public void AddExternalLinkBuilderTest()
        {
            Workbook workbook = new Workbook();
            ExternalLinkBuilder builder = new ExternalLinkBuilder(@"C:\Files\external.xlsx")
                .AddWorksheet("ExternalSheet");

            ExternalLink externalLink = workbook.AddExternalLink(builder);

            ExternalLink result = Assert.Single(workbook.GetExternalLinks());
            Assert.Same(externalLink, result);
            Assert.Same(builder.Build(), result);
            Assert.Equal("ExternalSheet", Assert.Single(result.Worksheets).Name);
        }

        [Fact(DisplayName = "Test of adding an external link using a null builder")]
        public void AddNullExternalLinkBuilderTest()
        {
            Workbook workbook = new Workbook();

            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(
                () => workbook.AddExternalLink((ExternalLinkBuilder)null));

            Assert.Equal("builder", exception.ParamName);
            Assert.Empty(workbook.GetExternalLinks());
        }

        [Fact(DisplayName = "Test of retrieving all external links")]
        public void GetExternalLinksTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink first = new ExternalLink(@"C:\Files\first.xlsx");
            ExternalLink second = new ExternalLink(@"C:\Files\second.xlsx");
            workbook.AddExternalLink(first);
            workbook.AddExternalLink(second);

            IReadOnlyList<ExternalLink> result = workbook.GetExternalLinks();

            Assert.Equal(2, result.Count);
            Assert.Same(first, result[0]);
            Assert.Same(second, result[1]);
        }

        [Fact(DisplayName = "Test of retrieving external links from a new workbook")]
        public void GetEmptyExternalLinksTest()
        {
            Workbook workbook = new Workbook();

            IReadOnlyList<ExternalLink> result = workbook.GetExternalLinks();

            Assert.NotNull(result);
            Assert.Empty(result);
        }

        [Fact(DisplayName = "Test of removing an existing external link")]
        public void RemoveExternalLinkTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink externalLink = new ExternalLink(@"C:\Files\external.xlsx");
            workbook.AddExternalLink(externalLink);

            bool result = workbook.RemoveExternalLink(externalLink);

            Assert.True(result);
            Assert.Empty(workbook.GetExternalLinks());
        }

        [Fact(DisplayName = "Test of removing an unrelated external link")]
        public void RemoveUnrelatedExternalLinkTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink existing = new ExternalLink(@"C:\Files\existing.xlsx");
            ExternalLink unrelated = new ExternalLink(@"C:\Files\unrelated.xlsx");
            workbook.AddExternalLink(existing);

            bool result = workbook.RemoveExternalLink(unrelated);

            Assert.False(result);
            ExternalLink remaining = Assert.Single(workbook.GetExternalLinks());
            Assert.Same(existing, remaining);
        }

        [Fact(DisplayName = "Test of removing one external link from two entries")]
        public void RemoveExternalLinkFromMultipleEntriesTest()
        {
            Workbook workbook = new Workbook();
            ExternalLink first = new ExternalLink(@"C:\Files\first.xlsx");
            ExternalLink second = new ExternalLink(@"C:\Files\second.xlsx");
            workbook.AddExternalLink(first);
            workbook.AddExternalLink(second);

            bool result = workbook.RemoveExternalLink(first);

            Assert.True(result);
            ExternalLink remaining = Assert.Single(workbook.GetExternalLinks());
            Assert.Same(second, remaining);
        }

        [Fact(DisplayName = "Test of clearing external links")]
        public void ClearExternalLinksTest()
        {
            Workbook workbook = new Workbook();
            workbook.AddExternalLink(new ExternalLink(@"C:\Files\first.xlsx"));
            workbook.AddExternalLink(new ExternalLink(@"C:\Files\second.xlsx"));

            workbook.ClearExternalLinks();

            Assert.Empty(workbook.GetExternalLinks());
        }

        [Fact(DisplayName = "Test of clearing external links from a new workbook")]
        public void ClearEmptyExternalLinksTest()
        {
            Workbook workbook = new Workbook();

            workbook.ClearExternalLinks();

            Assert.Empty(workbook.GetExternalLinks());
        }
    }
}

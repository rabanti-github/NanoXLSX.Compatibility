using NanoXLSX.Exceptions;
using NanoXLSX.Extensions;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class ExternalLinkWorkbookInlineWriterTest
    {

        [Fact(DisplayName = "Test of the unused properties for null (for coverage)")]
        public void NoOpPropertiesTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets", "definedNames");
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            Assert.Null(writer.XmlElement);
            Assert.Null(writer.WriteContext);
        }


        [Fact(DisplayName = "Test that workbook XML remains unchanged without external links")]
        public void LeavesWorkbookXmlUnchangedWithoutExternalLinks()
        {
            Workbook workbook = new Workbook("Sheet1");
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets", "definedNames");
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            writer.Execute();

            Assert.Equal(new[] { "sheets", "definedNames" }, GetChildNames(root));
        }

        [Fact(DisplayName = "Test writing workbook external references without defined names")]
        public void WritesExternalReferencesWithoutDefinedNames()
        {
            Workbook workbook = new Workbook("Sheet1");
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\first.xlsx"));
            StoreRelationshipId(workbook, 0, "rId7");
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets");
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            writer.Execute();

            Assert.Equal(new[] { "sheets", "externalReferences" }, GetChildNames(root));
        }

        [Fact(DisplayName = "Test writing ordered workbook external references and resolved defined names")]
        public void WritesExternalReferencesAndResolvedDefinedNames()
        {
            Workbook workbook = new Workbook("Sheet1");
            const string originalExpression = @"C:\data\[first.xlsx]Data!A1";
            DefinedName definedName = workbook.AddDefinedNameFormula("ExternalName", originalExpression);
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\first.xlsx"), 0);
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\second.xlsx"), 1);
            StoreRelationshipId(workbook, 0, "rId7");
            StoreRelationshipId(workbook, 1, "rId8");
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY,
                new Dictionary<int, ExternalLinkResolution>
                {
                    [0] = new ExternalLinkResolution("[1]Data!A1", new List<int> { 1 })
                });
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("fileVersion", "sheets", "definedNames", "calcPr");
            XmlElement definedNames = Assert.Single(root.FindChildElementsByName("definedNames"));
            XmlElement xmlDefinedName = XmlElement.CreateElement("definedName");
            xmlDefinedName.AddAttribute("name", "ExternalName");
            xmlDefinedName.InnerValue = originalExpression;
            definedNames.AddChildElement(xmlDefinedName);
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            writer.Execute();

            Assert.Equal(
                new[] { "fileVersion", "sheets", "externalReferences", "definedNames", "calcPr" },
                GetChildNames(root));
            XmlElement references = Assert.Single(root.FindChildElementsByName("externalReferences"));
            Assert.Collection(
                references.Children,
                item => Assert.Equal("rId7", ExternalLinkTestUtils.GetAttribute(item, "id")),
                item => Assert.Equal("rId8", ExternalLinkTestUtils.GetAttribute(item, "id")));
            Assert.Equal("[1]Data!A1", xmlDefinedName.InnerValue);
            Assert.Equal(originalExpression, definedName.TextValue);
        }

        [Theory(DisplayName = "Test ignoring unavailable defined-name resolution maps")]
        [InlineData(false)]
        [InlineData(true)]
        public void IgnoresUnavailableDefinedNameResolutions(bool storeEmptyMap)
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddDefinedNameFormula("ExternalName", @"C:\data\[first.xlsx]Data!A1");
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\first.xlsx"));
            StoreRelationshipId(workbook, 0, "rId7");
            if (storeEmptyMap)
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY,
                    new Dictionary<int, ExternalLinkResolution>());
            }
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets", "definedNames");
            XmlElement container = Assert.Single(root.FindChildElementsByName("definedNames"));
            XmlElement name = XmlElement.CreateElement("definedName");
            name.AddAttribute("name", "ExternalName");
            name.InnerValue = "unchanged";
            container.AddChildElement(name);
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            writer.Execute();

            Assert.Equal("unchanged", name.InnerValue);
        }

        [Fact(DisplayName = "Test ignoring a defined-name resolution with an unrelated index")]
        public void IgnoresUnrelatedDefinedNameResolution()
        {
            Workbook workbook = new Workbook("Sheet1");
            workbook.AddDefinedNameFormula("ExternalName", @"C:\data\[first.xlsx]Data!A1");
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\first.xlsx"));
            StoreRelationshipId(workbook, 0, "rId7");
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY,
                new Dictionary<int, ExternalLinkResolution>
                {
                    [4] = new ExternalLinkResolution("[1]Data!A1", new List<int> { 1 }) // ID = 4
                });
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets", "definedNames");
            XmlElement container = Assert.Single(root.FindChildElementsByName("definedNames"));
            XmlElement name = XmlElement.CreateElement("definedName");
            name.AddAttribute("name", "ExternalName");
            name.InnerValue = "unchanged";
            container.AddChildElement(name);
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            writer.Execute();

            Assert.Equal("unchanged", name.InnerValue);
        }

        [Fact(DisplayName = "Test failure when an external-link relationship ID is unavailable")]
        public void RejectsMissingExternalLinkRelationshipId()
        {
            Workbook workbook = new Workbook("Sheet1");
            ExternalLinkTestUtils.StoreExternalLink(workbook, new ExternalLink(@"C:\data\first.xlsx"));
            XmlElement root = ExternalLinkTestUtils.CreateWorkbookRoot("sheets");
            ExternalLinkWorkbookInlineWriter writer = CreateWriter(workbook, ref root);

            IOException exception = Assert.Throws<IOException>(() => writer.Execute());

            Assert.Contains(CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + "0", exception.Message);
        }

        private static ExternalLinkWorkbookInlineWriter CreateWriter(Workbook workbook, ref XmlElement root)
        {
            ExternalLinkWorkbookInlineWriter writer = new ExternalLinkWorkbookInlineWriter();
            writer.Init(ref root, workbook);
            return writer;
        }

        private static void StoreRelationshipId(Workbook workbook, int index, string relationshipId)
        {
            workbook.AuxiliaryData.SetData(
                PlugInUUID.WriterPackageRegistryQueue,
                PlugInUUID.PackagePartRelationshipId,
                CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + index,
                relationshipId);
        }

        private static string[] GetChildNames(XmlElement root)
        {
            return root.Children.ConvertAll(child => child.Name).ToArray();
        }
    }
}

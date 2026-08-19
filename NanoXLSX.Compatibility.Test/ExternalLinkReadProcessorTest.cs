using NanoXLSX.Internal;
using NanoXLSX.Internal.Reader;
using NanoXLSX.Registry;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Xunit;

namespace NanoXLSX.Compatibility.Test
{
    public class ExternalLinkReadProcessorTest
    {
        [Fact(DisplayName = "Test retention of relationship order and raw external defined names during workbook reading")]
        public void WorkbookReaderRetainsRelationshipOrderAndRawExternalDefinedNames()
        {
            const string xml =
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                "<externalReferences><externalReference r:id=\"rIdFirst\"/><externalReference r:id=\"rIdSecond\"/></externalReferences>" +
                "<definedNames><definedName name=\"ExternalCell\">'[1]Data'!$A$1</definedName>" +
                "<definedName name=\"LocalName\">Local!$A$1</definedName></definedNames></workbook>";
            Workbook workbook = new Workbook("Local");
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(xml)))
            {
                ExternalLinkWorkbookReader reader = new ExternalLinkWorkbookReader();
                reader.Init(stream, workbook, null);
                reader.Execute();
            }

            Assert.Equal(
                new[] { "rIdFirst", "rIdSecond" },
                workbook.AuxiliaryData.GetData<List<string>>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY));
            List<ExternalDefinedNameReference> definitions = workbook.AuxiliaryData.GetData<List<ExternalDefinedNameReference>>(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY);
            ExternalDefinedNameReference definition = Assert.Single(definitions);
            Assert.Equal("ExternalCell", definition.Name);
            Assert.Equal("'[1]Data'!$A$1", definition.Expression);
            Assert.Null(definition.LocalSheetIndex);
        }

        [Fact(DisplayName = "Test replacement of one-based external link IDs in formula defined names")]
        public void ReplacesOneBasedExternalLinkIdInFormulaDefinedName()
        {
            Workbook workbook = new Workbook("Local");
            workbook.AddDefinedNameFormula("ExternalTotal", "SUM('[1]Data'!$A$1:$A$2)");
            AddExternalLinkData(workbook,
                new[] { "rIdExternal1" },
                CreateExternalLink(@"C:\data\source.xlsx", "rIdExternal1"));

            Execute(workbook);

            DefinedName result = workbook.GetDefinedName("ExternalTotal");
            Assert.Equal("SUM('[C:\\data\\source.xlsx]Data'!$A$1:$A$2)", result.TextValue);
            Assert.DoesNotContain("[1]", result.TextValue);
        }

        [Fact(DisplayName = "Test reconstruction of external cell defined names as resolved formulas")]
        public void RebuildsExternalCellDefinedNameAsResolvedFormula()
        {
            Workbook workbook = new Workbook("Local");
            DefinedName original = DefinedName.ResolveDefinedName(
                "ExternalCell",
                "'[1]Data'!$A$1",
                workbook,
                null,
                "External reference");
            Assert.Equal(DefinedName.NameType.Cell, original.Type);
            workbook.AddDefinedName(original);
            workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY,
                new List<ExternalDefinedNameReference>
                {
                    new ExternalDefinedNameReference("ExternalCell", null, "'[1]Data'!$A$1")
                });
            AddExternalLinkData(workbook,
                new[] { "rIdExternal1" },
                CreateExternalLink(@"C:\data\source.xlsx", "rIdExternal1"));

            Execute(workbook);

            DefinedName result = workbook.GetDefinedName("ExternalCell");
            Assert.NotSame(original, result);
            Assert.Equal(DefinedName.NameType.Formula, result.Type);
            Assert.Equal("'[C:\\data\\source.xlsx]Data'!$A$1", result.TextValue);
            Assert.Equal("External reference", result.Comment);
            Assert.True(result.HasExternalReferences);
        }

        [Fact(DisplayName = "Test resolution of external link IDs by workbook relationship order")]
        public void ResolvesIdsByWorkbookRelationshipOrder()
        {
            Workbook workbook = new Workbook("Local");
            workbook.AddDefinedNameFormula(
                "ExternalTotal",
                "SUM('[1]First'!A1,'[2]Second'!A1)");
            ExternalLink first = CreateExternalLink(@"C:\data\first.xlsx", "rIdFirst");
            ExternalLink second = CreateExternalLink(@"C:\data\second.xlsx", "rIdSecond");
            AddExternalLinkData(
                workbook,
                new[] { "rIdFirst", "rIdSecond" },
                second,
                first);

            Execute(workbook);

            Assert.Equal(
                "SUM('[C:\\data\\first.xlsx]First'!A1,'[C:\\data\\second.xlsx]Second'!A1)",
                workbook.GetDefinedName("ExternalTotal").TextValue);
        }

        [Fact(DisplayName = "Test rejection of defined name IDs without matching external references")]
        public void RejectsDefinedNameIdWithoutMatchingExternalReference()
        {
            Workbook workbook = new Workbook("Local");
            workbook.AddDefinedNameFormula("ExternalTotal", "'[2]Data'!A1");
            AddExternalLinkData(workbook,
                new[] { "rIdExternal1" },
                CreateExternalLink(@"C:\data\source.xlsx", "rIdExternal1"));

            Assert.Throws<Exceptions.IOException>(() => Execute(workbook));
            Assert.Equal("'[2]Data'!A1", workbook.GetDefinedName("ExternalTotal").TextValue);
        }

        private static ExternalLink CreateExternalLink(string uri, string workbookRelationshipId)
        {
            ExternalLink link = new ExternalLink();
            link.SetReadUris(uri, null, null);
            link.WorkbookRId = workbookRelationshipId;
            return link;
        }

        private static void AddExternalLinkData(
            Workbook workbook,
            IReadOnlyList<string> relationshipIds,
            params ExternalLink[] externalLinks)
        {
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY,
                new List<string>(relationshipIds));
            for (int i = 0; i < externalLinks.Length; i++)
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.CompatibilityInlineProcessor,
                    CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                    i,
                    externalLinks[i],
                    true);
            }
        }

        private static void Execute(Workbook workbook)
        {
            ExternalLinkReadProcessor processor = new ExternalLinkReadProcessor();
            processor.Init(workbook, null);
            processor.Execute();
        }
    }
}

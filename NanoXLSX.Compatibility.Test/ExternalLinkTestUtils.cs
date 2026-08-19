using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Text;

namespace NanoXLSX.Compatibility.Test
{
    internal static class ExternalLinkTestUtils
    {
        internal const string ExternalLinkPathRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        internal const string ExternalLinkDocumentType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        internal const string ExternalLinkPartPath = "xl/externalLinks/externalLink1.xml";

        /// <summary>
        /// Creates a memory stream from a UTF-8 xml string
        /// </summary>
        internal static MemoryStream CreateStream(string xml)
        {
            return new MemoryStream(Encoding.UTF8.GetBytes(xml));
        }

        /// <summary>
        /// Gets the value of an XmlElement attribute, or returns null, if not found
        /// </summary>
        internal static string GetAttribute(XmlElement element, string name)
        {
            XmlAttribute? attribute = XmlAttribute.FindAttribute(name, element.Attributes);
            return attribute?.Value;
        }

        /// <summary>
        /// Creates the root XmlElement of a workbook document (rudimentary)
        /// </summary>
        internal static XmlElement CreateWorkbookRoot(params string[] childNames)
        {
            XmlElement root = XmlElement.CreateElement("workbook");
            foreach (string childName in childNames)
            {
                root.AddChildElement(XmlElement.CreateElement(childName));
            }
            return root;
        }

        /// <summary>
        /// Creates a relationship info object for an externalLink1.xml document
        /// </summary>
        internal static RelationshipInfo CreateCurrentRelationship()
        {
            return new RelationshipInfo(
                "rIdWorkbookExternal1",
                ExternalLinkDocumentType,
                "externalLinks/externalLink1.xml",
                TargetMode.Internal,
                "xl/_rels/workbook.xml.rels",
                "xl/workbook.xml",
                ExternalLinkPartPath);
        }

        /// <summary>
        /// Creates a relationship catalog, usually created on read by a discovery reader
        /// </summary>
        internal static RelationshipCatalog CreateRelationshipCatalog(
            string targetId = "rId1",
            string target = @"..\data\external.xlsx",
            string relationshipType = ExternalLinkPathRelationshipType,
            TargetMode targetMode = TargetMode.External)
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            catalog.TryAdd(new RelationshipInfo(
                targetId,
                relationshipType,
                target,
                targetMode,
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                ExternalLinkPartPath,
                null));
            return catalog;
        }

        /// <summary>
        /// Adds an additional relationship to the given relationship catalog
        /// </summary>
        internal static void AddExternalRelationship(
            RelationshipCatalog catalog,
            string id,
            string target,
            string relationshipType = ExternalLinkPathRelationshipType,
            TargetMode targetMode = TargetMode.External)
        {
            catalog.TryAdd(new RelationshipInfo(
                id,
                relationshipType,
                target,
                targetMode,
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                ExternalLinkPartPath,
                null));
        }

        /// <summary>
        /// Executes an instance of ExternalLinkReader on a new workbook with given XML content
        /// </summary>
        internal static Workbook ExecuteExternalLinkReader(string xml, RelationshipCatalog catalog)
        {
            Workbook workbook = new Workbook();
            if (catalog != null)
            {
                workbook.AuxiliaryData.SetData(
                    PlugInUUID.DiscoveryReader,
                    PlugInUUID.DiscoveryCatalogEntity,
                    catalog);
            }
            using (MemoryStream stream = CreateStream(xml))
            {
                ExternalLinkReader reader = new ExternalLinkReader
                {
                    CurrentRelationship = CreateCurrentRelationship()
                };
                reader.Init(stream, workbook, null, null);
                reader.Execute();
            }
            return workbook;
        }

        /// <summary>
        /// Executes an instance of ExternalLinkWorkbookReader on a new workbook with given XML content
        /// </summary>
        internal static Workbook ExecuteExternalLinkWorkbookReader(string xml)
        {
            Workbook workbook = new Workbook();
            using (MemoryStream stream = CreateStream(xml))
            {
                ExternalLinkWorkbookReader reader = new ExternalLinkWorkbookReader();
                reader.Init(stream, workbook, null);
                reader.Execute();
            }
            return workbook;
        }

        /// <summary>
        /// Creates an XML string of a external link document
        /// </summary>
        internal static string CreateExternalLinkXml(
            string targetRelationshipId = "rId1",
            string alternateUrls = null,
            string sheetNames = null,
            string definedNames = null,
            string sheetDataSet = null)
        {
            string targetAttribute = targetRelationshipId == null ? string.Empty : " r:id=\"" + targetRelationshipId + "\"";
            return
                "<externalLink xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:xxl21=\"http://schemas.microsoft.com/office/spreadsheetml/2021/extlinks2021\">" +
                "<externalBook" + targetAttribute + ">" +
                (alternateUrls ?? string.Empty) +
                (sheetNames ?? string.Empty) +
                (definedNames ?? string.Empty) +
                (sheetDataSet ?? string.Empty) +
                "</externalBook></externalLink>";
        }

        /// <summary>
        /// Creates an XML string of cached worksheet / cell data in a external link document
        /// </summary>
        internal static string CreateCachedCellXml(string cellType, string value, string attributes = null)
        {
            string typeAttribute = cellType == null ? string.Empty : " t=\"" + cellType + "\"";
            string valueElement = value == null ? string.Empty : "<v>" + value + "</v>";
            return CreateExternalLinkXml(
                sheetNames: "<sheetNames><sheetName val=\"Data\"/></sheetNames>",
                sheetDataSet:
                    "<sheetDataSet><sheetData sheetId=\"0\"><row r=\"1\"><cell r=\"A1\"" +
                    typeAttribute + (attributes ?? string.Empty) + ">" + valueElement +
                    "</cell></row></sheetData></sheetDataSet>");
        }

        /// <summary>
        /// Stores a given workbook as stream, loads and returns it
        /// </summary>
        internal static Workbook RoundTrip(Workbook workbook)
        {
            EnsureCompatibilityPluginsLoaded();
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                return WorkbookReader.Load(stream);
            }
        }

        [ExcludeFromCodeCoverage]
        private static void EnsureCompatibilityPluginsLoaded()
        {
            if (PlugInLoader.HasQueuePlugins(PlugInUUID.CompatibilityInlineProcessor) &&
                PlugInLoader.HasQueuePlugins(PlugInUUID.PreparingInlineProcessor) &&
                PlugInLoader.HasQueuePlugins(PlugInUUID.WriterPackageRegistryQueue) &&
                PlugInLoader.HasQueuePlugins(PlugInUUID.WriterAppendingQueue) &&
                PlugInLoader.HasQueuePlugins(PlugInUUID.ReaderPrependingQueue) &&
                PlugInLoader.HasQueuePlugins(PlugInUUID.FinalizingInlineProcessor))
            {
                return;
            }
            PlugInLoader.DisposePlugins();
            PlugInLoader.Initialize();
        }

        /// <summary>
        /// Mock class, representing a base writer, usually provided from a XlsxWriter process
        /// </summary>
        internal sealed class TestBaseWriter : IBaseWriter
        {
            private readonly HashSet<string> preparedFeatures = new HashSet<string>();

            public Workbook Workbook { get; }

            [ExcludeFromCodeCoverage]
            public IWriterProcessingData WriterProcessingData => null;

            [ExcludeFromCodeCoverage]
            public ISharedStringWriter SharedStringWriter { get; set; }

            internal TestBaseWriter(Workbook workbook)
            {
                Workbook = workbook;
            }

            [ExcludeFromCodeCoverage]
            public void MarkFeatureAsPrepared(string featureUuid)
            {
                preparedFeatures.Add(featureUuid);
            }

            [ExcludeFromCodeCoverage]
            public bool IsFeaturePrepared(string featureUuid)
            {
                return preparedFeatures.Contains(featureUuid);
            }
        }

        /// <summary>
        /// Stores an ExternalLink object into AuxiliaryData of the given workbook
        /// </summary>
        internal static void StoreExternalLink(Workbook workbook, ExternalLink link, int index = 0)
        {
            workbook.AuxiliaryData.SetData(
                PlugInUUID.CompatibilityInlineProcessor,
                CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY,
                index,
                link,
                true);
        }
    }
}

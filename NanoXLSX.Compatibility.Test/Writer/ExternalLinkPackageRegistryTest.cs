using NanoXLSX.Extensions;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Registry;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO.Packaging;
using System.Linq;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class ExternalLinkPackageRegistryTest
    {
        private const string PackagePartPath = "xl/externalLinks/";
        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        private const string RelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        private const string RelationshipPathType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        [Fact(DisplayName = "Test registration of package parts for two external links")]
        public void RegistersPackagePartsForTwoExternalLinks()
        {
            const int initialOrderNumber = 100;
            Workbook workbook = new Workbook("Sheet1");
            workbook.AuxiliaryData.SetData(
                PlugInUUID.WriterPackageRegistryQueue,
                PlugInUUID.LastPackageOrderNumber,
                initialOrderNumber);
            workbook.AddExternalLink(new ExternalLink(@"C:\data\first.xlsx"));
            workbook.AddExternalLink(new ExternalLink(@"C:\data\second.xlsx", @"..\data\second.xlsx"));
            ExternalLinkPackageRegistry registry = new ExternalLinkPackageRegistry();
            registry.Init(new TestBaseWriter(workbook));

            registry.Execute();

            const int expectedCount = 2;
            Assert.Equal(new[] { initialOrderNumber + 1, initialOrderNumber + 2 }, registry.OrderNumbers);
            Assert.Equal(expectedCount, registry.OrderNumbers.Distinct().Count());
            Assert.Equal(new[] { PackagePartPath, PackagePartPath }, registry.PackagePartPaths);
            Assert.Equal(new[] { "externalLink1.xml", "externalLink2.xml" }, registry.PackagePartFileNames);
            Assert.Equal(new[] { ContentType, ContentType }, registry.ContentTypes);
            Assert.Equal(new[] { RelationshipType, RelationshipType }, registry.RelationshipTypes);
            Assert.Equal(new[] { false, false }, registry.ArePackagePartsRoot);
            Assert.Equal(
                new[]
                {
                    CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + "0",
                    CompatibilityConstants.UNIQUE_PACKAGE_PART_INDEX_PREFIX + "1"
                },
                registry.UniquePackagePartIndices);

            Assert.Equal(expectedCount, registry.PackagePartPaths.Count);
            Assert.Equal(expectedCount, registry.PackagePartFileNames.Count);
            Assert.Equal(expectedCount, registry.ContentTypes.Count);
            Assert.Equal(expectedCount, registry.RelationshipTypes.Count);
            Assert.Equal(expectedCount, registry.ArePackagePartsRoot.Count);
            Assert.Equal(expectedCount, registry.UniquePackagePartIndices.Count);
            Assert.Equal(expectedCount, registry.PackagePartRelationships.Count);

            AssertNoNullOrEmptyEntries(registry.PackagePartPaths);
            AssertNoNullOrEmptyEntries(registry.PackagePartFileNames);
            AssertNoNullOrEmptyEntries(registry.ContentTypes);
            AssertNoNullOrEmptyEntries(registry.RelationshipTypes);
            AssertNoNullOrEmptyEntries(registry.UniquePackagePartIndices);

            Assert.Single(registry.PackagePartRelationships[0]);
            Assert.Equal(2, registry.PackagePartRelationships[1].Count);
            Assert.All(registry.PackagePartRelationships, relationships =>
            {
                Assert.NotNull(relationships);
                Assert.NotEmpty(relationships);
                Assert.All(relationships, relationship =>
                {
                    Assert.False(string.IsNullOrWhiteSpace(relationship.RelationshipId));
                    Assert.Equal(RelationshipPathType, relationship.RelationshipType);
                    Assert.False(string.IsNullOrWhiteSpace(relationship.Target));
                    Assert.Equal(TargetMode.External, relationship.TargetMode);
                });
            });
        }

        private static void AssertNoNullOrEmptyEntries(IReadOnlyList<string> entries)
        {
            Assert.All(entries, entry => Assert.False(string.IsNullOrWhiteSpace(entry)));
        }

        private sealed class TestBaseWriter : IBaseWriter
        {
            private readonly HashSet<string> preparedFeatures = new HashSet<string>();

            public Workbook Workbook { get; }

            [ExcludeFromCodeCoverage]
            public IWriterProcessingData WriterProcessingData => null;

            [ExcludeFromCodeCoverage]
            public ISharedStringWriter SharedStringWriter { get; set; }

            public TestBaseWriter(Workbook workbook)
            {
                Workbook = workbook;
            }

            public void MarkFeatureAsPrepared(string featureUuid)
            {
                preparedFeatures.Add(featureUuid);
            }

            public bool IsFeaturePrepared(string featureUuid)
            {
                return preparedFeatures.Contains(featureUuid);
            }
        }
    }
}

using NanoXLSX.Internal.Writer;
using System.IO.Packaging;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class PluginPackageRelationshipTest
    {

        [Fact(DisplayName = "Test of the properties assignment (for coverage)")]
        public void PropertiesTest()
        {
            PluginPackageRelationship rel = new PluginPackageRelationship();
            Assert.Null(rel.Target);
            Assert.Null(rel.RelationshipType);
            Assert.Null(rel.RelationshipId);
            Assert.Equal(TargetMode.Internal, rel.TargetMode); // Implicitly set

            rel.Target = "C:\\temp\\file";
            rel.RelationshipType = "https://schema.org/some_schema";
            rel.RelationshipId = "rId5";
            rel.TargetMode = TargetMode.External;

            Assert.Equal("C:\\temp\\file", rel.Target);
            Assert.Equal("https://schema.org/some_schema", rel.RelationshipType);
            Assert.Equal("rId5", rel.RelationshipId);
            Assert.Equal(TargetMode.External, rel.TargetMode);
        }

    }
}

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Internal.Writer;
using NanoXLSX.Registry;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using Xunit;

namespace NanoXLSX.Compatibility.Test.Writer
{
    public class CompatibilityInlineWriteProcessorTest
    {
        [Fact(DisplayName = "Test marking of the external link writer feature as prepared")]
        public void MarksExternalLinkFeatureAsPrepared()
        {
            TestWriteContext context = new TestWriteContext(new Workbook("Sheet1"));
            CompatibilityInlineWriteProcessor processor = new CompatibilityInlineWriteProcessor();
            processor.Init(context);

            processor.Execute();

            Assert.True(context.IsFeaturePrepared(PlugInUUID.WriteExternalLinkFeature));
        }

        private sealed class TestWriteContext : IWriteContext
        {
            private readonly HashSet<string> preparedFeatures = new HashSet<string>();

            [ExcludeFromCodeCoverage]
            public Workbook Workbook { get; }

            [ExcludeFromCodeCoverage]
            public IWriterProcessingData WriterProcessingData => null;

            public TestWriteContext(Workbook workbook)
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

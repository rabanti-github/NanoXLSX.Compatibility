/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class responsible to mark compatibility features as enabled, to be written to a XLSX file
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "MAIN_COMPATIBILITY_WRITE_INLINE_ENABLE_PROCESSOR", QueueUUID = PlugInUUID.CompatibilityInlineProcessor, PlugInOrder = 1000)]
    internal class CompatibilityInlineWriteProcessor : IPluginInlineWriteProcessor
    {
        /// <summary>
        /// Write context
        /// </summary>
        public IWriteContext WriteContext { get; set; }

        /// <summary>
        /// Initializing method
        /// </summary>
        /// <param name="context">Writ context</param>
        public void Init(IWriteContext context)
        {
            this.WriteContext = context;
        }

        /// <summary>
        /// Main execution method of the processor
        /// </summary>
        public void Execute()
        {
            // Tells the writer that external links can be written
            WriteContext.MarkFeatureAsPrepared(PlugInUUID.WriteExternalLinkFeature);

            // TODO Add further enabled compatibility features here
        }
    }
}

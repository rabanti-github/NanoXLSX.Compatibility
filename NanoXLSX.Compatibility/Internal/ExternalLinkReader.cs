using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class implementing a reader for exteranl link files of XLSX files.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_READER", QueueUUID = PlugInUUID.ReaderPrependingQueue, PlugInOrder = 20001)]
    internal class ExternalLinkReader : IPluginQueueReader
    {
        #region privateFields

        private readonly List<ExternalLink> externalLinks;
        private Stream stream;

        #endregion

        #region properties
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Current workbook
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reference to a ReaderPlugInHandler, to be used for prepending operations in the <see cref="Execute"/> method
        /// </summary>
        /// 
        /// Reference to a ReaderPlugInHandler, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public ExternalLinkReader()
        {
            externalLinks = new List<ExternalLink>();
        }
        #endregion

        #region methods
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">Stream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Inline plug-in handler</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                {

                }

            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }
        #endregion

        #region sub-classes

        #endregion
    }
}
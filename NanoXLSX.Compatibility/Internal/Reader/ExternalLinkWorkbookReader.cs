/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace NanoXLSX.Internal.Reader
{
    /// <summary>
    /// Class implementing a reader for external link references in workbook definitions.
    /// </summary>
    [NanoXlsxQueuePlugIn(PlugInUUID = "EXTERNAL_LINK_WORKBOOK_READER", QueueUUID = PlugInUUID.WorkbookInlineReader, PlugInOrder = 10000)]
    internal class ExternalLinkWorkbookReader : IPluginInlineReader
    {
        private Stream stream;
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
        public Workbook Workbook { get; set; }

        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, int? index = null)
        {
            this.stream = stream;
            this.Workbook = workbook;
        }

        public void Execute()
        {
            try
            {
                List<string> rIds = new List<string>();
                using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                {
                    while (reader.Read())
                    {
                        if (XmlStreamUtils.IsElement(reader, "externalReferences"))
                        {
                            ReadExternalReferencesRIds(reader.ReadSubtree(), rIds);
                            Workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY, rIds);
                            break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }

        private static void ReadExternalReferencesRIds(XmlReader references, List<string> rIds)
        {
            while (references.Read())
            {
                if (XmlStreamUtils.IsElement(references, "externalReference"))
                {
                    string rid = references.GetAttribute("r:id");
                    if (rid != null)
                    {
                        rIds.Add(rid);
                    }
                }
            }
        }
    }
}

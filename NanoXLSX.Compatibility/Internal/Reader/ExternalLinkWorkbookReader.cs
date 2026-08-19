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
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
                List<ExternalDefinedNameReference> definedNames = new List<ExternalDefinedNameReference>();
                using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                {
                    while (reader.Read())
                    {
                        if (XmlStreamUtils.IsElement(reader, "externalReferences"))
                        {
                            ReadExternalReferencesRIds(reader.ReadSubtree(), rIds);
                        }
                        else if (XmlStreamUtils.IsElement(reader, "definedNames"))
                        {
                            ReadExternalDefinedNames(reader.ReadSubtree(), definedNames);
                        }
                    }
                }
                if (rIds.Count > 0)
                {
                    Workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY, rIds);
                }
                if (definedNames.Count > 0)
                {
                    Workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_REFERENCE_DEFINED_NAMES_ENTITY, definedNames);
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

        private static void ReadExternalDefinedNames(XmlReader definitions, List<ExternalDefinedNameReference> definedNames)
        {
            while (definitions.Read())
            {
                if (!XmlStreamUtils.IsElement(definitions, "definedName"))
                {
                    continue;
                }

                string name = definitions.GetAttribute("name");
                string localSheetId = definitions.GetAttribute("localSheetId");
                int? localSheetIndex = string.IsNullOrEmpty(localSheetId)
                    ? (int?)null
                    : ParserUtils.ParseInt(localSheetId);
                string expression = ReadElementText(definitions);
                if (ExternalLinkFormulaUtils.DetectExternalLinkId(expression))
                {
                    definedNames.Add(new ExternalDefinedNameReference(name, localSheetIndex, expression));
                }
            }
        }

        private static string ReadElementText(XmlReader reader)
        {
            if (reader.IsEmptyElement)
            {
                return string.Empty;
            }
            StringBuilder builder = new StringBuilder();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    break;
                }
                if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA ||
                    reader.NodeType == XmlNodeType.SignificantWhitespace)
                {
                    builder.Append(reader.Value);
                }
            }
            return builder.ToString();
        }
    }

    /// <summary>
    /// Raw external defined-name data retained until external-link relationships can be resolved.
    /// </summary>
    internal sealed class ExternalDefinedNameReference
    {
        public string Name { get; }
        public int? LocalSheetIndex { get; }
        public string Expression { get; }

        public ExternalDefinedNameReference(string name, int? localSheetIndex, string expression)
        {
            Name = name;
            LocalSheetIndex = localSheetIndex;
            Expression = expression;
        }
    }
}

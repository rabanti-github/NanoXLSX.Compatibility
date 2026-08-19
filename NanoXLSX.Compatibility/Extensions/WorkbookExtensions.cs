/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Internal;
using NanoXLSX.Registry;
using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Writer extension methods for the <see cref="Workbook">Workbook</see> class
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class WorkbookExtensions
    {
        /// <summary>
        /// Adds an external link to the workbook.
        /// </summary>
        /// <param name="workbook">Workbook instance</param>
        /// <param name="externalLink">External link object to add</param>
        /// <exception cref="ArgumentNullException">Thrown if the external link object was null</exception>
        public static void AddExternalLink(this Workbook workbook, ExternalLink externalLink)
        {
            if (externalLink == null)
            {
                throw new ArgumentNullException(nameof(externalLink), "An external link cannot be null.");
            }
            IReadOnlyList<ExternalLink> externalLinks = GetExternalLinks(workbook);
            int index = externalLinks.Count;
            while (workbook.AuxiliaryData.GetData<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY, index) != null)
            {
                index++;
            }
            workbook.AuxiliaryData.SetData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY, index, externalLink, true);
        }

        /// <summary>
        /// Adds an external link to the workbook, using a builder.
        /// </summary>
        /// <param name="workbook">Workbook instance</param>
        /// <param name="builder"></param>
        /// <returns>Returns the created and added external link</returns>
        /// <exception cref="ArgumentNullException">Thrown if the builder was null</exception>
        /// \remark <remarks>The builder has to be populated first before passing to this method</remarks>
        public static ExternalLink AddExternalLink(this Workbook workbook, ExternalLinkBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder), "A builder for external link cannot be null.");
            }
            ExternalLink externalLink = builder.Build();
            AddExternalLink(workbook, externalLink);
            return externalLink;
        }

        /// <summary>
        /// Gets all external links of the workbook.
        /// </summary>
        /// <param name="workbook">Workbook instance</param>
        /// <returns>Returns a read-only list of external links, or an empty list, if none are defined in the workbook</returns>
        public static IReadOnlyList<ExternalLink> GetExternalLinks(this Workbook workbook)
        {
            return workbook.AuxiliaryData.GetDataList<ExternalLink>(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
        }

        /// <summary>
        /// Removes a specific external link from the workbook
        /// </summary>
        /// <param name="workbook">Workbook instance</param>
        /// <param name="externalLink">Object to remove</param>
        /// <returns>Returns true if successfully deleted, otherwise false</returns>
        public static bool RemoveExternalLink(this Workbook workbook, ExternalLink externalLink)
        {
            return workbook.AuxiliaryData.RemoveEntityData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY, externalLink);
        }

        /// <summary>
        /// Removes all external links from the workbook
        /// </summary>
        /// <param name="workbook">Workbook instance</param>
        public static void ClearExternalLinks(this Workbook workbook)
        {
            workbook.AuxiliaryData.ClearEntityData(PlugInUUID.CompatibilityInlineProcessor, CompatibilityConstants.EXTERNAL_LINK_OBJECT_ENTITY);
        }

    }
}

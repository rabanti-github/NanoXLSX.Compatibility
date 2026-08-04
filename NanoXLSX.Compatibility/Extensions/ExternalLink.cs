/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a link to an external workbook.
    /// </summary>
    public class ExternalLink
    {

        private readonly List<ExternalWorksheet> worksheets = new List<ExternalWorksheet>();
        private readonly List<ExternalDefinedName> definedNames = new List<ExternalDefinedName>();
        private readonly List<string> uris = new List<string>();

        /// <summary>
        /// Gets the paths or URI of the external workbook.
        /// </summary>
        public IReadOnlyList<string> Uris => uris;

        /// <summary>
        /// Gets the worksheets declared for the external workbook.
        /// The worksheet index corresponds to the sheetId used in OOXML.
        /// </summary>
        public IReadOnlyList<ExternalWorksheet> Worksheets => worksheets;

        /// <summary>
        /// Gets the defined names declared by the external workbook.
        /// </summary>
        public IReadOnlyList<ExternalDefinedName> DefinedNames => definedNames;

        /// <summary>
        /// Constructor of an external workbook link.
        /// </summary>
        public ExternalLink()
        {
        }

        /// <summary>
        /// Constructor of an external workbook link with URI.
        /// </summary>
        /// <param name="uri">Main path or URI of the external workbook.</param>
        /// \remark <remarks>The URI is often defined as absolute path. For a better portability, a relative path can be used. However, NanoXLSX will not access or validate the defined URI</remarks>
        public ExternalLink(string uri)
        {
            AddUri(uri);
        }

        /// <summary>
        /// Creates a builder for this external workbook link.
        /// </summary>
        public ExternalLinkBuilder CreateBuilder()
        {
            return new ExternalLinkBuilder(this);
        }

        /// <summary>
        /// Adds a defined name from the external workbook.
        /// </summary>
        /// <param name="name">Name of the external defined name.</param>
        /// <param name="refersTo">Expression referenced by the defined name, for example ='Sheet1'!$B$1.</param>
        public void AddDefinedName(string name, string refersTo)
        {
            AddDefinedName(new ExternalDefinedName(name, refersTo));
        }

        /// <summary>
        /// Adds a defined name from the external workbook.
        /// </summary>
        /// <param name="definedName">Defined name object</param>
        public void AddDefinedName(ExternalDefinedName definedName)
        {
            if (definedName == null)
            {
                throw new ArgumentException("An external defined name cannot be null");
            }

            if (definedNames.Any(
                    item => string.Equals(
                        item.Name,
                        definedName.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"The external defined name '{definedName.Name}' already exists.");
            }
            definedNames.Add(definedName);
            //return this;
        }

        /// <summary>
        /// Adds a URI of an external workbook
        /// </summary>
        /// <param name="uri">URI of the external workbook</param>
        internal void AddUri(string uri)
        {
            if (string.IsNullOrWhiteSpace(uri))
            {
                throw new ArgumentException("The URI cannot be null or empty");
            }
            uris.Add(uri);
        }

        /// <summary>
        /// Adds an external worksheet
        /// </summary>
        /// <param name="worksheet">External Worksheet to Add</param>
        internal void AddWorksheet(ExternalWorksheet worksheet)
        {
            if (worksheet == null)
            {
                throw new ArgumentException("The worksheet cannot be null");
            }

            if (worksheets.Any(
                    item => string.Equals(
                        item.Name,
                        worksheet.Name,
                        StringComparison.OrdinalIgnoreCase)))
            {
                throw new ArgumentException($"The external worksheet '{worksheet.Name}' already exists.");
            }
            worksheets.Add(worksheet);
        }

        /// <summary>
        /// Gets an external worksheet by name.
        /// </summary>
        /// <param name="name">Name of the external defined name</param>
        public ExternalWorksheet GetWorksheet(string name)
        {
            ExternalWorksheet worksheet = worksheets.FirstOrDefault(
                item => string.Equals(
                    item.Name,
                    name,
                    StringComparison.OrdinalIgnoreCase));

            if (worksheet == null)
            {
                throw new ArgumentException($"The external worksheet '{name}' does not exist.");
            }

            return worksheet;
        }

    }
}

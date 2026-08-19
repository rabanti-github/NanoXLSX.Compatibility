/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using static NanoXLSX.ExternalCellValue;

namespace NanoXLSX
{
    /// <summary>
    /// Provides a fluent API for defining worksheets and cached cell values
    /// of an external workbook link.
    /// </summary>
    public class ExternalLinkBuilder
    {
        private readonly ExternalLink externalLink;
        private ExternalWorksheet currentWorksheet;

        /// <summary>
        /// Constructor with mandatory absolute URI and optional relative URI 
        /// </summary>
        /// <param name="absoluteUri">Absolute URI</param>
        /// <param name="relativeUri">Relative URI, default is null</param>
        public ExternalLinkBuilder(string absoluteUri, string relativeUri = null)
        {
            ExternalLink externalLink = new ExternalLink(absoluteUri, relativeUri);
            this.externalLink = externalLink;
        }

        /// <summary>
        /// Internal constructor of the builder
        /// </summary>
        /// <param name="externalLink"></param>
        /// <exception cref="ArgumentException"></exception>
        internal ExternalLinkBuilder(ExternalLink externalLink)
        {
            this.externalLink = externalLink ?? throw new ArgumentException("The external link cannot be null");
        }

        /// <summary>
        /// Returns the configured external workbook link.
        /// </summary>
        /// <returns>Returns the external link object with added worksheets, defined names and cells</returns>
        public ExternalLink Build()
        {
            return externalLink;
        }

        /// <summary>
        /// Adds a worksheet to the external workbook definition and marks it as currently used worksheet in the builder.
        /// as the current worksheet.
        /// </summary>
        /// <exception cref="Exceptions.FormatException">Thrown if the given name is not a valid worksheet name</exception>
        public ExternalLinkBuilder AddWorksheet(string name)
        {
            ExternalWorksheet worksheet = new ExternalWorksheet(name);

            externalLink.AddWorksheet(worksheet);
            currentWorksheet = worksheet;

            return this;
        }

        /// <summary>
        /// Selects an already declared external worksheet. The worksheet will be marked as currently used worksheet in the builder.
        /// </summary>
        /// <param name="name">Name of the external worksheet to use</param>
        public ExternalLinkBuilder UseWorksheet(string name)
        {
            currentWorksheet = externalLink.GetWorksheet(name);
            return this;
        }

        /// <summary>
        /// Adds or replaces a cached cell value (string representation) on the current worksheet.
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="value">
        /// Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case
        /// </param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        public ExternalLinkBuilder AddCell(string address, string value)
        {
            EnsureCurrentWorksheet();
            currentWorksheet.AddCell(address, value);

            return this;
        }

        /// <summary>
        /// Adds or replaces a cached cell value with specified type on the current worksheet.
        /// </summary>
        /// <param name="address">Address of the cell</param>
        /// <param name="value">
        /// Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case
        /// </param>
        /// <param name="type">Type of the external cell</param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        public ExternalLinkBuilder AddCell(string address, string value, ExternalCellValue.DataType type)
        {
            EnsureCurrentWorksheet();
            currentWorksheet.AddCell(address, value, type);
            return this;
        }

        /// <summary>
        /// Checks whether an external worksheet was added or selected for the builder
        /// </summary>
        private void EnsureCurrentWorksheet()
        {
            if (currentWorksheet == null)
            {
                throw new ArgumentException(
                    "No external worksheet is currently selected. " +
                    "Call AddWorksheet or UseWorksheet before adding cell data.");
            }
        }

    }
}

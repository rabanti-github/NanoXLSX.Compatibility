/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using static NanoXLSX.ExternalCellValue;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a worksheet declared by an external workbook link.
    /// </summary>
    public class ExternalWorksheet
    {

        /// <summary>
        /// Dictionary of the external cells (addresses with values as strings)
        /// </summary>
        private readonly Dictionary<Address, ExternalCellValue> cells = new Dictionary<Address, ExternalCellValue>();


        /// <summary>
        /// Gets the dictionary on external cells, where the <see cref="Address"/> is the key and <see cref="ExternalCellValue"/> is the value
        /// </summary>
        public ReadOnlyDictionary<Address, ExternalCellValue> Cells
        {
            get { return new ReadOnlyDictionary<Address, ExternalCellValue>(cells); }
        }

        /// <summary>
        /// Gets the name of the external worksheet.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Determines whether the values could not be refreshed. Currently just for roundtrip preservation on read and write
        /// </summary>
        internal bool? RefreshErros { get; set; }

        /// <summary>
        /// Constructor with name 
        /// </summary>
        /// <param name="name">Name of the external Worksheet</param>
        public ExternalWorksheet(string name)
        {
            Validators.ValidateWorksheetName(name);
            Name = name;
        }

        /// <summary>
        /// Adds or replaces a cached cell value.
        /// </summary>
        /// <param name="address">Address of the external cell</param>
        /// <param name="value">Sting representation of the external cell value, where the default <see cref="ExternalCellValue.DataType.String"/> is used. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case.</param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        public ExternalWorksheet AddCell(string address, string value)
        {
            return AddCell(address, value, ExternalCellValue.DataType.String);
        }

        /// <summary>
        /// Adds or replaces a cached cell value with defined typ.
        /// </summary>
        /// <param name="address">Address of the external cell</param>
        /// <param name="value">
        /// Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case
        /// </param>
        /// <param name="type">Type of the external cell</param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        public ExternalWorksheet AddCell(string address, string value, ExternalCellValue.DataType type)
        {
            Validators.ValidateCellAddressExpression(address, Cell.AddressScope.SingleAddress);
            cells[new Address(address)] = new ExternalCellValue(value, type);
            return this;

        }

        /// <summary>
        /// Internal method to add a cached value, provided by the reader (roundtrip) 
        /// </summary>
        /// <param name="address">Address of the external cell</param>
        /// <param name="value">
        /// Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case
        /// </param>
        /// <param name="type">Type of the external cell</param>
        /// <param name="cellMetaData">Optional metadata ID</param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        internal void AddCell(string address, string value, ExternalCellValue.DataType type, string cellMetaData)
        {
            AddCell(address, value, type);
            if (cellMetaData != null)
            {
                cells[new Address(address)].CellMetadata = ParserUtils.ParseInt(cellMetaData);
            }
        }

        /// <summary>
        /// Removes cached data for a cell.
        /// </summary>
        /// <param name="address">Address to remove</param>
        public bool RemoveCell(string address)
        {
            Validators.ValidateCellAddressExpression(address, Cell.AddressScope.SingleAddress);
            return cells.Remove(new Address(address));
        }

        /// <summary>
        /// Tries to get cached data for a cell.
        /// </summary>
        /// <param name="address">Address of the cached cell</param>
        /// <param name="cell">Out parameter of the cell value</param>
        public bool TryGetCell(string address, out ExternalCellValue cell)
        {
            try
            {
                return cells.TryGetValue(new Address(address), out cell);
            }
            catch (Exception)
            {
                cell = null;
                return false;
            }
        }
    }
}

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


        internal ExternalWorksheet(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("The external worksheet name must not be null or empty.");
            }
            Name = name;
        }

        /// <summary>
        /// Adds or replaces a cached cell value.
        /// </summary>
        /// <param name="address">Address of the external cell</param>
        /// <param name="value">Sting representation of the external cell value, where the default <see cref="ExternalCellValue.DataType.SharedString"/> is used. A null value will be transformed to non-cached value (<see cref="ExternalCellValue.DataType.Empty"/>), represented by an empty sting</param>
        public ExternalWorksheet AddCell(string address, string value)
        {
            return AddCell(address, value, ExternalCellValue.DataType.SharedString);
        }

        /// <summary>
        /// Adds or replaces a cached cell value with defined typ.
        /// </summary>
        /// <param name="address">Address of the external cell</param>
        /// <param name="value">Sting representation of the external cell value. A null value will be transformed to non-cached value (<see cref="ExternalCellValue.DataType.Empty"/>), represented by an empty string</param>
        /// <param name="type">Data type of the external, cached cell</param>
        public ExternalWorksheet AddCell(string address, string value, ExternalCellValue.DataType type)
        {
            Validators.ValidateCellAddressExpression(address, Cell.AddressScope.SingleAddress);
            cells[new Address(address)] = new ExternalCellValue(value, type);
            return this;

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

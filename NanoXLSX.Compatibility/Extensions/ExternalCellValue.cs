/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a cached cell of an external worksheet with an optional type
    /// </summary>
    public class ExternalCellValue
    {
        /// <summary>
        /// Enum for the data type of the external cell
        /// </summary>
        public enum DataType
        {
            /// <summary>External cell value is a number. This is the implicit default, if not specified</summary>
            Number,
            /// <summary>External cell value is boolean</summary>
            Boolean,
            /// <summary>External cell value is date or date and time (ISO 8601)</summary>
            Date,
            /// <summary>External cell value is an error and not actually a value</summary>
            Error,
            /// <summary>External cell value is a string reference</summary>
#pragma warning disable CA1720
            String,
#pragma warning restore CA1720
            /// <summary>Not a real type for external cells, but used to mark a non-cached cell value. "0" will be shown as cached value in a local worksheet</summary>
            Empty
        }

        /// <summary>
        /// Value of the external, cached cell as string representation
        /// </summary>
        public string Value { get; private set; }
        /// <summary>
        /// Type of the external, cached cell value
        /// </summary>
        /// \Remark <remarks>String values in external cell caches are commonly represented by Excel as <see cref="DataType.String"/>, even if the referenced source cell contains a formula</remarks>

        public DataType Type { get; private set; }

        /// <summary>
        /// Cell metadata ID. Currently just for roundtrip preservation on read and write
        /// </summary>
        internal int? CellMetadata { get; set; }

        /// <summary>
        /// Constructor with value and type
        /// </summary>
        /// <param name="value">
        /// Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case
        /// </param>
        /// <param name="type">Type of the external cell</param>
        /// <exception cref="ArgumentException">Thrown if a boolean value is not represented as "0" or "1".</exception>
        /// \Remark <remarks>The validity of numeric values is not checked. Boolean values must be represented as "0" or "1". The type <see cref="DataType.Empty"/> discards the passed value.</remarks>
        public ExternalCellValue(string value, DataType type)
        {
            if (value == null || type == DataType.Empty)
            {
                Value = "";
                Type = DataType.Empty;
                return;
            }

            Value = NormalizeValue(value, type);
            Type = type;
        }

        /// <summary>
        /// Constructor with value. The type <see cref="DataType.String"/> will be used as default type
        /// </summary>
        /// <param name="value">Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be an empty string in this case</param>
        /// \Remark <remarks>The validity of the passed string representation of a number is not checked.</remarks>
        public ExternalCellValue(string value) : this(value, DataType.String) { }

        /// <summary>
        /// Normalizes a cached value to its compliant representation
        /// </summary>
        /// <param name="value">Raw value</param>
        /// <param name="type">Type of the value</param>
        /// <returns></returns>
        /// <exception cref="ArgumentException">Thrown if a boolean is in a invalid form</exception>
        private static string NormalizeValue(string value, DataType type)
        {
            switch (type)
            {
                case DataType.Boolean:
                    if (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase))
                    {
                        return "1";
                    }
                    if (value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase))
                    {
                        return "0";
                    }
                    throw new ArgumentException("An invalid boolean value was provided (0 or 1 are expected): " + value);
                default:
                    return value;
            }
        }
    }
}

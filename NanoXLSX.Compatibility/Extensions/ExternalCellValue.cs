/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

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
            /// <summary>External cell value is an inline string with optional formatting, but not maintained as shared string</summary>
            InlineString,
            /// <summary>External cell value is a shared string reference</summary>
            SharedString,
            /// <summary>External cell value is a formula string</summary>
            Formula,
            /// <summary>Not a real type for external cells, but used to mark a non-cached cell value. "0" will be shown as cached value in a local worksheet</summary>
            Empty
        }

        /// <summary>
        /// Value of the external, cached cell as string representation
        /// </summary>
        public string Value { get; private set; }
        /// <summary>
        /// Type of the external, cached cell value. If not specified, the default <see cref="DataType.Number"/> will be used
        /// </summary>
        /// \Remark <remarks>References pointing to string values in cells are often specified by Excel as <see cref="DataType.Formula"/> although no actual formula is in place</remarks>

        public DataType Type { get; private set; }

        /// <summary>
        /// Constructor with value and type
        /// </summary>
        /// <param name="value">Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be "0" in this case</param>
        /// <param name="type">Type of the external cell</param>
        /// \Remark <remarks>The validity of the passed string representation of a number is not checked. The type <see cref="DataType.Empty"/> will discard the passed value</remarks>
        public ExternalCellValue(string value, DataType type)
        {
            if (value == null || type == DataType.Empty)
            {
                Value = "0";
                Type = DataType.Empty;
            }
            else
            {
                Value = value;
                Type = type;
            }
        }

        /// <summary>
        /// Constructor with value. The type <see cref="DataType.Number"/> will be used as default type
        /// </summary>
        /// <param name="value">Value as string representation. Null will be transformed to the type <see cref="DataType.Empty"/>. The cached value will be "0" case</param>
        /// \Remark <remarks>The validity of the passed string representation of a number is not checked.</remarks>
        public ExternalCellValue(string value) : this(value, DataType.Number) { }
    }
}

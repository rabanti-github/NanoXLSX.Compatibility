/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Represents a defined name declared by an external workbook.
    /// </summary>
    public class ExternalDefinedName
    {
        /// <summary>
        /// Gets the name of the external defined name.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the referenced expression.
        /// </summary>
        public string RefersTo { get; private set; }

        /// <summary>
        /// Constructor of an external defined name
        /// </summary>
        /// <param name="name">Name of the external defined name</param>
        /// <param name="refersTo">Referenced expression</param>
        /// <exception cref="FormatException">Thrown if the name or refersTo was null, empty or only white spaces</exception>
        public ExternalDefinedName(string name, string refersTo)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new FormatException("The defined name must not be null or empty.");
            }
            if (string.IsNullOrWhiteSpace(refersTo))
            {
                throw new FormatException("The defined name reference must not be null or empty.");
            }
            Name = name;
            RefersTo = refersTo;
        }
    }
}

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
        /// Relationship ID. Currently just for roundtrip preservation on read and write
        /// </summary>
        internal string RelationshipId { get; set; }

        /// <summary>
        /// Constructor of an external defined name
        /// </summary>
        /// <param name="name">Name of the external defined name</param>
        /// <param name="refersTo">Referenced expression</param>
        /// <exception cref="FormatException">Thrown if the name or refersTo was null, empty or only white spaces</exception>
        public ExternalDefinedName(string name, string refersTo) : this(name, refersTo, true)
        {
        }

        /// <summary>
        /// Internal constructor of an external defined name with the option to skip validation of refersTo (my be optional from the reader)
        /// </summary>
        /// <param name="name">Name of the external defined name</param>
        /// <param name="refersTo">Referenced expression</param>
        /// <param name="validateRefersTo">If true, the refersTo expression is checked (basic), otherwise unchecked</param>
        /// <exception cref="FormatException">Thrown if the name was null, empty or only white spaces. Same applies to refersTo if validateRefersTo is set to true</exception>
        internal ExternalDefinedName(string name, string refersTo, bool validateRefersTo)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new FormatException("The defined name must not be null or empty.");
            }
            if (validateRefersTo && string.IsNullOrWhiteSpace(refersTo))
            {
                throw new FormatException("The defined name reference must not be null or empty.");
            }
            Name = name;
            RefersTo = refersTo;
        }

    }
}

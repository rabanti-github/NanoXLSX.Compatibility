/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Internal class for reader and writer constants (e.g. for auxiliary data)
    /// </summary>
    internal class CompatibilityConstants
    {
        /// <summary>
        /// Entity ID for external link objects
        /// </summary>
        public const string EXTERNAL_LINK_OBJECT_ENTITY = "external-link-object-entity";

        /// <summary>
        /// Entity ID for resolved defined names
        /// </summary>
        public const string EXTERNAL_LINK_RESOLVED_DEFINED_NAMES_ENTITY = "external-link-resolved-defined-names-entity";

        /// <summary>
        /// Entity ID for resolved formulas
        /// </summary>
        public const string EXTERNAL_LINK_RESOLVED_FORMULAS_ENTITY = "external-link-resolved-formulas-entity";

        /// <summary>
        /// Entity ID for external reference workbook relationship IDs
        /// </summary>
        public const string EXTERNAL_REFERENCE_WORKBOOK_RID_ENTITY = "external-reference-workbook-rid-entity";
    }
}

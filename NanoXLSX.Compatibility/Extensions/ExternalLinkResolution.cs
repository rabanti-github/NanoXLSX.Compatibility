/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;

namespace NanoXLSX.Extensions
{
    /// <summary>
    /// Helper class, holding information about successfully resolved external links
    /// </summary>
    internal sealed class ExternalLinkResolution
    {
        /// <summary>
        /// Resolved expression
        /// </summary>
        public string Expression { get; }
        /// <summary>
        /// Indices (OOXML identifiers) of the external link part
        /// </summary>
        public List<int> LinkIndexes { get; }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="expression">Resolved expression</param>
        /// <param name="linkIndexes">Indices (OOXML identifiers) of the external link part</param>
        public ExternalLinkResolution(string expression, List<int> linkIndexes)
        {
            Expression = expression;
            LinkIndexes = linkIndexes;
        }
    }
}

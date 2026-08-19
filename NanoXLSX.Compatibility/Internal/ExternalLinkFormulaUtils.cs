/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Lightweight helpers for detecting and replacing external-link tokens in formulas and defined-name expressions.
    /// </summary>
    internal static class ExternalLinkFormulaUtils
    {
        /// <summary>
        /// Determines whether an expression contains any internal external workbook identifier such as "[1]".
        /// </summary>
        internal static bool DetectExternalLinkId(string expression)
        {
            return DetectExternalLinkId(expression, null);
        }

        /// <summary>
        /// Determines whether an expression contains an internal external workbook identifier such as "[1]".
        /// Identifiers inside Excel string constants and structured references are ignored.
        /// </summary>
        internal static bool DetectExternalLinkId(string expression, string targetId)
        {
            if (targetId != null && !ParserUtils.IsValidExternalLinkId(targetId))
            {
                throw new ArgumentException(
                    $"The target ID '{targetId}' is not a valid external link ID.",
                    nameof(targetId));
            }

            if (string.IsNullOrEmpty(expression))
            {
                return false;
            }

            int firstOpeningBracket = expression.IndexOf('[');
            if (firstOpeningBracket < 0)
            {
                return false;
            }

            if (targetId != null && expression.IndexOf(targetId, StringComparison.Ordinal) < 0)
            {
                return false;
            }

            int firstQuote = expression.IndexOf('"');
            int scanStart = firstQuote >= 0 && firstQuote < firstOpeningBracket
                ? firstQuote
                : firstOpeningBracket;
            bool insideStringConstant = false;

            for (int i = scanStart; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current == '"')
                {
                    if (insideStringConstant && i + 1 < expression.Length && expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    insideStringConstant = !insideStringConstant;
                    continue;
                }

                if (insideStringConstant || current != '[' ||
                    !ParserUtils.TryReadExternalLinkId(expression, i, out int identifierLength))
                {
                    continue;
                }

                if (targetId == null ||
                    (identifierLength == targetId.Length &&
                     string.CompareOrdinal(expression, i, targetId, 0, identifierLength) == 0))
                {
                    return true;
                }

                i += identifierLength - 1;
            }

            return false;
        }

        /// <summary>
        /// Finds a human-readable external workbook token which remains unresolved in an expression.
        /// Numeric identifiers, Excel string constants, and structured references are ignored.
        /// </summary>
        internal static bool TryFindUnresolvedExternalLink(string expression, out string token)
        {
            token = null;
            if (string.IsNullOrEmpty(expression) || expression.IndexOf('[') < 0)
            {
                return false;
            }
            bool insideStringConstant = false;
            for (int i = 0; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current == '"')
                {
                    if (insideStringConstant && i + 1 < expression.Length && expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    insideStringConstant = !insideStringConstant;
                    continue;
                }
                if (insideStringConstant || current != '[')
                {
                    continue;
                }
                int closingBracket = expression.IndexOf(']', i + 1);
                if (closingBracket <= i + 1)
                {
                    continue;
                }
                string bracketToken = expression.Substring(i, closingBracket - i + 1);
                if (ParserUtils.IsValidExternalLinkId(bracketToken) || HasStructuredReferencePrefix(expression, i))
                {
                    i = closingBracket;
                    continue;
                }
                if (FormsExternalReference(expression, closingBracket))
                {
                    token = bracketToken;
                    return true;
                }
                i = closingBracket;
            }

            return false;
        }

        /// <summary>
        /// Replaces internal external workbook identifiers with their preferred human-readable references.
        /// Unknown identifiers and identifiers inside Excel string constants are left unchanged.
        /// </summary>
        internal static string ReplaceExternalLinkId(string expression, Dictionary<string, ExternalLink> links)
        {
            if (expression == null || expression.Length == 0 || links == null || links.Count == 0)
            {
                return expression;
            }
            int firstOpeningBracket = expression.IndexOf('[');
            if (firstOpeningBracket < 0)
            {
                return expression;
            }
            bool insideStringConstant = false;
            int unchangedSectionStart = 0;
            StringBuilder builder = null;

            for (int i = firstOpeningBracket; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current == '"')
                {
                    if (insideStringConstant && i + 1 < expression.Length && expression[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }
                    insideStringConstant = !insideStringConstant;
                    continue;
                }

                if (insideStringConstant || current != '[' ||
                    !ParserUtils.TryReadExternalLinkId(expression, i, out int identifierLength))
                {
                    continue;
                }

                string identifier = expression.Substring(i, identifierLength);
                if (!links.TryGetValue(identifier, out ExternalLink externalLink))
                {
                    continue;
                }
                if (externalLink == null)
                {
                    throw new ArgumentException($"The external link '{identifier}' is null.", nameof(links));
                }

                string replacement = externalLink.ReadableReferenceToken;
                if (string.IsNullOrEmpty(replacement))
                {
                    throw new ArgumentException(
                        $"The external link '{identifier}' does not provide a usable URI.",
                        nameof(links));
                }

                if (builder == null)
                {
                    builder = new StringBuilder(expression.Length + replacement.Length);
                }
                builder.Append(expression, unchangedSectionStart, i - unchangedSectionStart);
                builder.Append(replacement);

                i += identifierLength - 1;
                unchangedSectionStart = i + 1;
            }

            if (builder == null)
            {
                return expression;
            }
            builder.Append(expression, unchangedSectionStart, expression.Length - unchangedSectionStart);
            return builder.ToString();
        }

        private static bool HasStructuredReferencePrefix(string expression, int openingBracket)
        {
            if (openingBracket == 0)
            {
                return false;
            }
            char previous = expression[openingBracket - 1];
            return char.IsLetterOrDigit(previous) || previous == '_' || previous == '.';
        }

        private static bool FormsExternalReference(string expression, int closingBracket)
        {
            bool hasWorksheetName = false;
            for (int i = closingBracket + 1; i < expression.Length; i++)
            {
                char current = expression[i];
                if (current == '!')
                {
                    return hasWorksheetName || i == closingBracket + 1;
                }
                if (IsReferenceTerminator(current))
                {
                    return false;
                }
                if (!char.IsWhiteSpace(current) && current != '\'')
                {
                    hasWorksheetName = true;
                }
            }
            return false;
        }

        private static bool IsReferenceTerminator(char character)
        {
            switch (character)
            {
                case '[':
                case ']':
                case '"':
                case '+':
                case '-':
                case '*':
                case '/':
                case '^':
                case '&':
                case '=':
                case '<':
                case '>':
                case ',':
                case ';':
                case '(':
                case ')':
                case '{':
                case '}':
                    return true;
                default:
                    return false;
            }
        }
    }
}

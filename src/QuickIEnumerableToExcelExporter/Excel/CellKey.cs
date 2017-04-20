/*
 * MIT License
 *
 * Copyright(c) 2017 Jann Westermann
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

using System;

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// A unique key of an excel cell within a worksheet
    /// </summary>
    internal class CellKey : IEquatable<CellKey>
    {
        /// <summary>
        /// Initializes a new instance of the class
        /// </summary>
        /// <param name="row">The 1-based index of the row</param>
        /// <param name="column">The 1-based index of the column</param>
        public CellKey(int row, int column)
        {
            Column = column;
            Row = row;
        }

        /// <summary>
        /// Gets the index of the column (1-based)
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the index of the row (1-based)
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Calculates the hash of this instance
        /// </summary>
        /// <returns>Hash of the current instance</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                var result = 0;
                result = (result * 397) ^ Column;
                result = (result * 397) ^ Row;

                return result;
            }
        }

        /// <summary>
        /// Compares this instance to another
        /// </summary>
        /// <param name="obj">The other instance</param>
        /// <returns>True if the keys are identical</returns>
        public override bool Equals(object obj) => obj != null && GetType() == obj.GetType() && Equals(obj as CellKey);

        /// <summary>
        /// Compares this instance to another
        /// </summary>
        /// <param name="other">The other instance</param>
        /// <returns>True if the keys are identical</returns>
        public bool Equals(CellKey other) => other != null && other.Column == Column && other.Row == Row;
    }
}
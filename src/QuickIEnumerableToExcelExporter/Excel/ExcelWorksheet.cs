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

using System.Collections.Generic;

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Represents an excel worksheet
    /// </summary>
    internal class ExcelWorksheet
    {
        /// <summary>
        /// Collection of cells within this worksheet
        /// </summary>
        private readonly Dictionary<CellKey, ExcelCell> _cells;

        /// <summary>
        /// Initializes a new instance of the class
        /// </summary>
        /// <param name="name">The name of the worksheet</param>
        public ExcelWorksheet(string name)
        {
            _cells = new Dictionary<CellKey, ExcelCell>();
            Name = name;
        }

        /// <summary>
        /// Gets the name of the worksheet
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the cell at the given position.
        /// </summary>
        /// <param name="row">The 1-based index of the row.</param>
        /// <param name="column">The 1-based index of the column.</param>
        /// <returns>The cell at this position</returns>
        public ExcelCell this[int row, int column]
        {
            get
            {
                var key = new CellKey(row, column);

                if (_cells.ContainsKey(key))
                {
                    return _cells[key];
                }

                var cell = new ExcelCell(row, column);
                _cells.Add(key, cell);

                return cell;
            }
        }
    }
}
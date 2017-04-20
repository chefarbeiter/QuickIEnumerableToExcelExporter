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

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Represents an cell within a excel worksheet
    /// </summary>
    internal class ExcelCell
    {
        /// <summary>
        /// Initializes a new instance of the class
        /// </summary>
        /// <param name="row">The 1-based index of the row</param>
        /// <param name="column">The 1-based index of the column</param>
        public ExcelCell(int row, int column)
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
        /// Gets or sets the value of the cell
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Gets or sets the format string used to display the value
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// Indicates if the font weight of this cell is bold
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// Indicates if the value is of type string
        /// </summary>
        public bool IsString => Value is string;
    }
}
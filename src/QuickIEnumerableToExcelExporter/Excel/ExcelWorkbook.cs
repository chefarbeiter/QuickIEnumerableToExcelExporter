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
using System.Collections.Generic;

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Represents an excel workbook (xlsx)
    /// </summary>
    internal class ExcelWorkbook
    {
        /// <summary>
        /// List of all worksheets within this workbokk
        /// </summary>
        private readonly List<ExcelWorksheet> _worksheets;

        /// <summary>
        /// Initializes a new instance of the class
        /// </summary>
        public ExcelWorkbook()
        {
            _worksheets = new List<Excel.ExcelWorksheet>();
        }

        /// <summary>
        /// Adds a new worksheet with the given name to the workbook
        /// </summary>
        /// <param name="name">Name of the worksheet</param>
        /// <returns>The created worksheet</returns>
        public ExcelWorksheet AddWorksheet(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));

            var worksheet = new ExcelWorksheet(name);
            _worksheets.Add(worksheet);

            return worksheet;
        }

        /// <summary>
        /// Writes the content of the workbook to the given file
        /// </summary>
        /// <param name="target">The target path</param>
        public void Save(string target)
        {
        }
    }
}
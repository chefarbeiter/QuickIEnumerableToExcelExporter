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
using System.Linq;

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Exports data to excel (xlsx)
    /// </summary>
    internal class ExcelExporter : IExporter
    {
        /// <summary>
        /// Path to the result file
        /// </summary>
        public string TargetPath { get; set; }

        /// <summary>
        /// The values to be exported (intermediate format)
        /// </summary>
        public IEnumerable<ExportRow> Rows { get; set; }

        /// <summary>
        /// Writes the data to excel
        /// </summary>
        public void Export()
        {
            if (string.IsNullOrEmpty(TargetPath)) throw new ArgumentNullException(nameof(TargetPath));
            if (Rows == null) throw new ArgumentNullException(nameof(Rows));

            var workbook = new ExcelWorkbook("Export");
            var rowsCount = Rows.Count();

            for (var rowIndex = 1; rowIndex <= rowsCount; rowIndex++)
            {
                var row = Rows.ElementAt(rowIndex - 1);

                for (var columnIndex = 1; columnIndex <= row.Values.Count; columnIndex++)
                {
                    var column = row.Values[columnIndex - 1];
                    var cell = workbook.Worksheet[rowIndex, columnIndex];

                    cell.Value = column.Value;
                    cell.IsBold = row.IsHeaderRow;
                    cell.Format = column.Format;
                }
            }

            workbook.Save(TargetPath);
        }
    }
}
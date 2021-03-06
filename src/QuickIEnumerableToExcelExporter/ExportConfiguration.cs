﻿/*
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

namespace QuickIEnumerableToExcelExporter
{
    /// <summary>
    /// Set of properties for configuring the behavior of the excel export.
    /// </summary>
    public class ExportConfiguration
    {
        /// <summary>
        /// If set to true, NULL values will be exported as string "NULL", instead of being ignored.
        /// The default value is false.
        /// </summary>
        public bool ExportNullAsString { get; set; } = false;

        /// <summary>
        /// If set to true, a row with headers will be exported.
        /// The default value is true.
        /// </summary>
        public bool WriteHeaderRow { get; set; } = true;
    }
}
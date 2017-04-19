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

namespace QuickIEnumerableToExcelExporter
{
    /// <summary>
    /// Marks an property to be exported to excel with special configuration.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportToExcelAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the header of the column for this property.
        /// Default is the name of the property.
        /// </summary>
        public string Header { get; set; }

        /// <summary>
        /// If set to true, this property will not be exported.
        /// Default is false.
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// Gets or sets a path to a property on the child object to serve as the export value of the object.
        /// </summary>
        public string DisplayMemberPath { get; set; }

        /// <summary>
        /// Gets or sets a string that specifies how to format the value
        /// </summary>
        public string StringFormat { get; set; }
    }
}
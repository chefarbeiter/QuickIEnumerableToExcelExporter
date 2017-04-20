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
using QuickIEnumerableToExcelExporter.Excel;

namespace QuickIEnumerableToExcelExporter
{
    /// <summary>
    ///
    /// </summary>
    public static class Extension
    {
        /// <summary>
        ///
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable"></param>
        /// <param name="target"></param>
        public static void ExportToExcel<T>(this IEnumerable<T> enumerable, string target)
            => ExportToExcel(enumerable, target, new ExportConfiguration());

        /// <summary>
        ///
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable"></param>
        /// <param name="target"></param>
        /// <param name="configuration"></param>
        public static void ExportToExcel<T>(this IEnumerable<T> enumerable, string target, ExportConfiguration configuration)
            => Export<T, ExcelExporter>(enumerable, target, configuration);

        /// <summary>
        ///
        /// </summary>
        /// <typeparam name="TItem"></typeparam>
        /// <typeparam name="TExporter"></typeparam>
        /// <param name="enumerable"></param>
        /// <param name="target"></param>
        /// <param name="configuration"></param>
        private static void Export<TItem, TExporter>(IEnumerable<TItem> enumerable, string target, ExportConfiguration configuration)
            where TExporter : IExporter, new()
        {
            // get metadata
            var metadata = MetadataReader.ReadMetadata(typeof(TItem));

            // write to intermediate
            var valuesReader = new ValuesReader<TItem>(metadata, configuration);
            var values = valuesReader.ReadValues(enumerable);

            // export intermediate to target
            var exporter = new TExporter()
            {
                TargetPath = target,
                Rows = values
            };

            exporter.Export();
        }
    }
}
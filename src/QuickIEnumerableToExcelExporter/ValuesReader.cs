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
using System.Linq;

namespace QuickIEnumerableToExcelExporter
{
    internal class ValuesReader<T>
    {
        private readonly List<ExportProperty> _properties;

        private readonly ExportConfiguration _configuration;

        public ValuesReader(List<ExportProperty> properties, ExportConfiguration configuration)
        {
            _properties = properties;
            _configuration = configuration;
        }

        public List<ExportRow> ReadValues(IEnumerable<T> enumerable)
        {
            var rows = new List<ExportRow>();

            if (_configuration.WriteHeaderRow)
            {
                rows.Add(new ExportRow
                {
                    IsHeaderRow = true,
                    Values = _properties.Select(p => new ExportValue { Value = p.Header }).ToList()
                });
            }

            rows.AddRange(enumerable.Select(item => ReadValuesForItem(item)));

            return rows;
        }

        private ExportRow ReadValuesForItem(T item)
        {
            var values = new List<ExportValue>();

            foreach (var property in _properties)
            {
                var currentProperty = property.Property;
                var value = currentProperty.PropertyInfo.GetValue(item);

                while (currentProperty.InnerProperty != null && value != null)
                {
                    currentProperty = currentProperty.InnerProperty;
                    value = currentProperty.PropertyInfo.GetValue(value);
                }

                if (value == null && _configuration.ExportNullAsString)
                {
                    value = "NULL";
                }

                values.Add(new ExportValue { Format = property.Format, Value = value });
            }

            return new ExportRow { Values = values };
        }
    }
}
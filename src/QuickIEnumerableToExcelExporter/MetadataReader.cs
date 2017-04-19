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

namespace QuickIEnumerableToExcelExporter
{
    internal static class MetadataReader
    {
        public static List<ExportProperty> ReadMetadata(Type exportType)
        {
            // Prepare output
            var exportProperties = new List<ExportProperty>();

            // Get properties and prepare
            exportType.GetProperties().ToList().ForEach(property =>
            {
                // Create with defaults
                var exportProperty = new ExportProperty
                {
                    Header = property.Name,
                    Property = new ExportPropertyInfo(property)
                };

                // Apply settings from attribute if given
                var exportAttribute = property.GetCustomAttributes(typeof(ExportToExcelAttribute), false).FirstOrDefault() as ExportToExcelAttribute;
                if (exportAttribute != null && !exportAttribute.Ignore)
                {
                    ApplyExportAttributeToProperty(exportAttribute, exportProperty);
                }

                // Add to collection if not ignored
                if (exportAttribute == null || !exportAttribute.Ignore)
                {
                    exportProperties.Add(exportProperty);
                }
            });

            return exportProperties;
        }

        private static void ApplyExportAttributeToProperty(ExportToExcelAttribute attribute, ExportProperty property)
        {
            if (!string.IsNullOrEmpty(attribute.Header))
            {
                property.Header = attribute.Header;
            }

            if (!string.IsNullOrEmpty(attribute.StringFormat))
            {
                property.Format = attribute.StringFormat;
            }

            if (!string.IsNullOrEmpty(attribute.DisplayMemberPath))
            {
                var parts = attribute.DisplayMemberPath.Split('.');
                var currentParentProperty = property.Property;

                for (var partIndex = 0; partIndex < parts.Length; partIndex++)
                {
                    var part = parts[partIndex];
                    var partProperty = currentParentProperty.PropertyInfo.PropertyType.GetProperty(part);
                    var partExportProperty = new ExportPropertyInfo(partProperty);

                    currentParentProperty.InnerProperty = partExportProperty;
                    currentParentProperty = partExportProperty;
                }
            }
        }
    }
}
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
using System.Reflection;

namespace QuickIEnumerableToExcelExporter
{
    internal class ExportPropertyInfo
    {
        public PropertyInfo Property { get; set; }

        public ExportPropertyInfo InnerProperty { get; set; }
    }
}
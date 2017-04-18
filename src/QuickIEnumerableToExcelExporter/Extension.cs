using System.Collections.Generic;

namespace QuickIEnumerableToExcelExporter
{
    public static class Extension
    {
        public static void ExportToExcel<T>(this IEnumerable<T> enumerable, string target)
            => ExportToExcel(enumerable, target, new ExportConfiguration());

        public static void ExportToExcel<T>(this IEnumerable<T> enumerable, string target, ExportConfiguration configuration)
        {
            // get metadata

            // write to intermediate

            // convert intermediate to excel
        }
    }
}
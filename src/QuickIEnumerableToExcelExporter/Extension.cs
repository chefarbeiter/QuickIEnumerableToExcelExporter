using System.Collections.Generic;

namespace QuickIEnumerableToExcelExporter
{
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
        {
            // get metadata

            // write to intermediate

            // convert intermediate to excel
        }
    }
}
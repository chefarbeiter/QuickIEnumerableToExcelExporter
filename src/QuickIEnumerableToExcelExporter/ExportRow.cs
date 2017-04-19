using System.Collections.Generic;

namespace QuickIEnumerableToExcelExporter
{
    /// <summary>
    /// Describes an row for the export
    /// </summary>
    internal class ExportRow
    {
        /// <summary>
        /// Indicates if the current row is a header
        /// </summary>
        public bool IsHeaderRow { get; set; }

        /// <summary>
        /// List of all values (columns)
        /// </summary>
        public List<ExportValue> Values { get; set; }
    }
}
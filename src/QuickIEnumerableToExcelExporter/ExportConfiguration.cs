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
    }
}
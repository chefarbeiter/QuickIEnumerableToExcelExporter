namespace QuickIEnumerableToExcelExporter
{
    /// <summary>
    /// Describes an value of the export
    /// </summary>
    internal class ExportValue
    {
        /// <summary>
        /// The value to be exported
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// String with the format for displaying the value
        /// </summary>
        public string Format { get; set; }
    }
}
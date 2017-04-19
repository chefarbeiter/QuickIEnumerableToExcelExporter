using System;
using QuickIEnumerableToExcelExporter;

namespace UnitTests
{
    internal class TestExportItem
    {
        public string Name { get; set; }

        [ExportToExcel(Header = "Identity", StringFormat = "#,#0")]
        public int Id { get; set; }

        [ExportToExcel(DisplayMemberPath = "Year")]
        public DateTime Timestamp { get; set; }

        [ExportToExcel(DisplayMemberPath = "Inner.Value")]
        public NestedItem Nested { get; set; }
    }

    internal class NestedItem
    {
        public string Value { get; set; }

        public NestedItem Inner { get; set; }
    }

    internal class TestExportItemWithIgnore
    {
        public int Id { get; set; }

        public string Name { get; set; }

        [ExportToExcel(Ignore = true)]
        public string Password { get; set; }
    }
}
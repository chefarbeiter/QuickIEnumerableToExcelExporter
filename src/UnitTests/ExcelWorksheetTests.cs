using QuickIEnumerableToExcelExporter.Excel;
using Xunit;

namespace UnitTests
{
    public class ExcelWorksheetTests
    {
        [Fact]
        public void SetAndReadCellValue()
        {
            var worksheet = new ExcelWorksheet("test");
            worksheet[10, 10].Value = "test_value";

            var value = worksheet[10, 10].Value;

            Assert.Equal("test_value", value);
        }
    }
}
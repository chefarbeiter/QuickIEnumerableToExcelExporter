using System;
using QuickIEnumerableToExcelExporter.Excel;
using Xunit;

namespace UnitTests
{
    public class ExcelWorkbookTests
    {
        [Fact]
        public void AddWorksheet()
        {
            var workbook = new ExcelWorkbook("test");
            var worksheet = workbook.Worksheet;

            Assert.NotNull(worksheet);
            Assert.Equal("test", worksheet.Name);
        }

        [Fact]
        public void AddWorksheetWithoutName()
        {
            Assert.Throws<ArgumentNullException>(() => new ExcelWorkbook(null));
            Assert.Throws<ArgumentNullException>(() => new ExcelWorkbook(String.Empty));
        }
    }
}
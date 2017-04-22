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
            var workbook = new ExcelWorkbook();
            var worksheet = workbook.AddWorksheet("test");

            Assert.NotNull(worksheet);
            Assert.Equal("test", worksheet.Name);
        }

        [Fact]
        public void AddWorksheetWithoutName()
        {
            var workbook = new ExcelWorkbook();

            Assert.Throws<ArgumentNullException>(() => workbook.AddWorksheet(null));
            Assert.Throws<ArgumentNullException>(() => workbook.AddWorksheet(String.Empty));
        }
    }
}
using System.Collections.Generic;
using System.IO;
using QuickIEnumerableToExcelExporter;
using QuickIEnumerableToExcelExporter.Excel;
using Xunit;

namespace UnitTests
{
    public class ExcelExporterTests
    {
        [Fact]
        public void ExportFileShouldBeCreated()
        {
            var target = Path.GetTempFileName();
            File.Delete(target);

            var exporter = new ExcelExporter
            {
                Rows = new List<ExportRow>(),
                TargetPath = target
            };

            exporter.Export();

            Assert.True(File.Exists(target));

            File.Delete(target);
        }
    }
}
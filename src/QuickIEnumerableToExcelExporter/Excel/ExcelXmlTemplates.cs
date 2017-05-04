namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Contains XML templates
    /// </summary>
    internal static class ExcelXmlTemplates
    {
        /// <summary>
        /// Template of the workbook xml.
        /// Parameter {0} = The name of the worksheet
        /// </summary>
        public const string Workbook =
            @"<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheets><sheet name=""{0}"" sheetId=""1"" r:id=""rId1""/></sheets></workbook>";

        /// <summary>
        /// Template of a worksheet
        /// Parameters:
        /// {0} - Address of last filled cell
        /// {1} - The collection of rows xml
        /// </summary>
        public const string Worksheet =
            @"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><dimension ref=""A1:{0}""/><sheetFormatPr defaultRowHeight=""15"" /><cols><col min=""5"" max=""5"" width=""10.140625"" bestFit=""1"" customWidth=""1""/></cols><sheetData>{1}</sheetData></worksheet>";

        /// <summary>
        /// Template of a row
        /// Parameters:
        /// {0} - The 1-based index of the row
        /// {1} - The 1-based index of the last column
        /// {2} - The cells xml
        /// </summary>
        public const string WorksheetRow = @"<row r=""{0}"" spans=""1:{1}"">{2}</row>";

        /// <summary>
        /// Template of a cell
        /// {0} - The address of the cell
        /// {1} - The type of the cell value
        /// {2} - The cell value
        /// </summary>
        public const string WorksheetCell = @"<c r=""{0}"" t=""{1}""><v>{2}</v></c>";

        /// <summary>
        /// Template of the sharedStrings file.
        /// Parameters:
        /// {0} - Count and Unique Count
        /// {1} - Strings
        /// </summary>
        public const string SharedStrings =
            @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""{0}"" uniqueCount=""{0}"">{1}</sst>";

        /// <summary>
        /// Single shared string.
        /// Parameter {0} = The String
        /// </summary>
        public const string SharedStringXml = @"<si><t>{0}</t></si>";
    }
}
/*
 * MIT License
 *
 * Copyright(c) 2017 Jann Westermann
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 *
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;

namespace QuickIEnumerableToExcelExporter.Excel
{
    /// <summary>
    /// Represents an excel workbook (xlsx)
    /// </summary>
    internal class ExcelWorkbook
    {
        /// <summary>
        /// Initializes a new workbook with a worksheet
        /// </summary>
        /// <param name="sheetName">Name of the sheet</param>
        public ExcelWorkbook(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName)) throw new ArgumentNullException(nameof(sheetName));

            Worksheet = new ExcelWorksheet(sheetName);
        }

        /// <summary>
        /// The worksheet with the data
        /// </summary>
        public ExcelWorksheet Worksheet { get; }

        /// <summary>
        /// Writes the content of the workbook to the given file
        /// </summary>
        /// <param name="target">The target path</param>
        public void Save(string target)
        {
            using (var package = Package.Open(target, FileMode.Create))
            {
                var workbook = package.CreatePart(new Uri("/xl/workbook.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                var worksheet = package.CreatePart(new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                var styles = package.CreatePart(new Uri("/xl/styles.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                var sharedStringsPart = package.CreatePart(new Uri("/xl/sharedStrings.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");

                package.CreateRelationship(workbook.Uri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1");
                workbook.CreateRelationship(new Uri("worksheets/sheet1.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", "rId1");
                workbook.CreateRelationship(new Uri("sharedStrings.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", "rId3");
                workbook.CreateRelationship(new Uri("styles.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "rId2");

                CreateSharedStringsPart(sharedStringsPart);
                CreateWorkbookPart(workbook);
                CreateWorksheetPart(worksheet);

                using (var writer = new StreamWriter(styles.GetStream()))
                {
                    writer.WriteLine(stylesXml);
                }
            }
        }

        /// <summary>
        /// Creates the sharedStrings.xml part
        /// </summary>
        /// <param name="sharedStringsFile">The created sharedStrings.xml package part</param>
        private void CreateSharedStringsPart(PackagePart sharedStringsFile)
        {
            // create shared strings collection
            var sharedStrings = new List<string>();
            var stringValues = Worksheet.Cells.Where(c => c.Value.IsString).Select(c => c.Value).ToList();

            foreach (var stringValue in stringValues)
            {
                var value = stringValue.Value.ToString();

                if (sharedStrings.Contains(value))
                {
                    var index = sharedStrings.IndexOf(value) + 1;
                    stringValue.Value = index.ToString();
                }
                else
                {
                    sharedStrings.Add(value);
                    stringValue.Value = sharedStrings.Count.ToString();
                }
            }

            // create xml
            var sharedStringsXml = string.Join(string.Empty, sharedStrings.Select(str => string.Format(ExcelXmlTemplates.SharedStringXml, str)));
            var xml = string.Format(ExcelXmlTemplates.SharedStrings, sharedStrings.Count, sharedStringsXml);

            // write xml
            using (var writer = new StreamWriter(sharedStringsFile.GetStream()))
            {
                writer.Write(xml);
            }
        }

        /// <summary>
        /// Writes the content to the workbook.xml file
        /// </summary>
        /// <param name="workbookFile">The created workbook package part</param>
        private void CreateWorkbookPart(PackagePart workbookFile)
        {
            using (var writer = new StreamWriter(workbookFile.GetStream()))
            {
                writer.Write(string.Format(ExcelXmlTemplates.Workbook, Worksheet.Name));
            }
        }

        private void CreateWorksheetPart(PackagePart worksheetFile)
        {
            //            var cellsByRows = Worksheet.Cells.GroupBy(c => c.Key.Row).Select( )

            using (var writer = new StreamWriter(worksheetFile.GetStream()))
            {
                writer.Write(string.Empty);
            }
        }

        private const string stylesXml = @"<styleSheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
  <numFmts count=""0"" />
  <fonts count=""1"">
    <font>
      <sz val=""11"" />
      <name val=""Calibri"" />
    </font>
  </fonts>
  <fills count=""2"">
    <fill>
      <patternFill patternType=""none"" />
    </fill>
    <fill>
      <patternFill patternType=""gray125"" />
    </fill>
  </fills>
  <borders count=""1"">
    <border>
      <left />
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
  </borders>
  <cellStyleXfs count=""1"">
    <xf numFmtId=""0"" fontId=""0"" />
  </cellStyleXfs>
  <cellXfs count=""1"">
    <xf numFmtId=""0"" applyNumberFormat=""1"" fontId=""0"" applyFont=""1"" xfId=""0"" />
  </cellXfs>
  <cellStyles count=""1"">
    <cellStyle name=""Normal"" xfId=""0"" builtinId=""0"" />
  </cellStyles>
  <dxfs count=""0"" />
</styleSheet>";
    }
}
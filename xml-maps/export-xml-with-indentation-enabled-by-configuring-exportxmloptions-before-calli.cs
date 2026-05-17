using System;
using Aspose.Cells;

namespace AsposeCellsExportXmlWithIndentation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Id");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Configure save options to enable indented XML formatting
            SpreadsheetML2003SaveOptions saveOptions = new SpreadsheetML2003SaveOptions
            {
                IsIndentedFormatting = true   // Child elements will be indented
            };

            // Export the workbook as an indented XML file
            workbook.Save("ExportedIndented.xml", saveOptions);

            Console.WriteLine("XML exported with indentation to 'ExportedIndented.xml'.");
        }
    }
}
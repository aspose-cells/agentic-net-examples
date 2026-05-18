using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["C1"].PutValue("City");

            // Populate some data rows
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["C2"].PutValue("New York");

            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["C3"].PutValue("London");

            // Configure JSON save options to include header row
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                HasHeaderRow = true,          // column headers will be used as JSON keys
                ExportEmptyCells = false,     // optional: do not export empty cells
                Indent = "  "                 // optional: pretty‑print with two spaces
            };

            // Save the workbook as JSON using the options (lifecycle rule: save)
            workbook.Save("WorkbookExport.json", jsonOptions);
        }
    }
}
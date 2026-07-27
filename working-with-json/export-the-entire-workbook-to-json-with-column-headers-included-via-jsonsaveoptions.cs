using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("City");

        // Add sample data rows
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["C2"].PutValue("New York");

        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["C3"].PutValue("London");

        // Configure JSON save options to include column headers
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            HasHeaderRow = true,          // treat first row as header
            ExportEmptyCells = false,     // omit empty cells
            Indent = "  "                 // pretty‑print with two spaces
        };

        // Export the entire workbook to JSON using the configured options
        workbook.Save("WorkbookExport.json", jsonOptions);
    }
}
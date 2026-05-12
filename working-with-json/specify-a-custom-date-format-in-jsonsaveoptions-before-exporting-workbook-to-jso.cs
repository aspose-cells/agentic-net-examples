using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data including a date value
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("BirthDate");
            sheet.Cells["A2"].PutValue("John Doe");
            sheet.Cells["B2"].PutValue(new DateTime(1990, 5, 15));

            // Apply a custom date format to the date cell (B2)
            // This format will be used when the workbook is saved to JSON
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MM-yyyy";   // Custom date format
            sheet.Cells["B2"].SetStyle(dateStyle);

            // Configure JSON save options (no specific date format property exists,
            // the cell's style determines how dates are represented)
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                ExportAsString = false,          // Export native types
                ExportEmptyCells = false,
                HasHeaderRow = true,
                ExportNestedStructure = false,
                Indent = "    "                  // Optional: pretty‑print JSON
            };

            // Save the workbook as a JSON file using the configured options
            string outputPath = "ExportedData.json";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to JSON with custom date format at: {outputPath}");
        }
    }
}
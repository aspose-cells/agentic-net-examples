using System;
using Aspose.Cells;
using Aspose.Cells.Json;

namespace AsposeCellsJsonExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data including a DateTime value
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("BirthDate");
            sheet.Cells["A2"].PutValue("John Doe");
            sheet.Cells["B2"].PutValue(new DateTime(1990, 5, 15));

            // Apply a custom date format to the cell that contains the date
            // This format will be used when the cell value is exported as a string
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MM-yyyy";   // Custom date format
            sheet.Cells["B2"].SetStyle(dateStyle);

            // Configure JSON save options
            JsonSaveOptions saveOptions = new JsonSaveOptions
            {
                // Export cell values as strings so that the custom date format is preserved
                ExportAsString = true,

                // Optional: set indentation for pretty‑printed JSON
                Indent = "    "   // 4 spaces
            };

            // Save the workbook as a JSON file using the configured options
            string outputPath = "ExportedWithCustomDateFormat.json";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to JSON with custom date format at: {outputPath}");
        }
    }
}
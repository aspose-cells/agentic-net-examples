using System;
using Aspose.Cells;

namespace AsposeCellsJsonExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data with an empty row (row index 2)
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");
            // Row 3 (index 2) left empty intentionally
            sheet.Cells["A4"].PutValue("Data3");
            sheet.Cells["B4"].PutValue("Data4");

            // Configure JSON save options to skip empty rows
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // When SkipEmptyRows is true, empty rows are omitted from the output
                SkipEmptyRows = true
            };

            // Save the workbook as JSON using the configured options
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"JSON saved to '{outputPath}' with empty rows excluded.");
        }
    }
}
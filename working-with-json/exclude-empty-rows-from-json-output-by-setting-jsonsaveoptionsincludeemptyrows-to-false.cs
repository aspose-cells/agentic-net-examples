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

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with an empty row (row index 2)
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue("Data1");
            cells["B2"].PutValue("Data2");
            // Row 3 (index 2) left empty intentionally
            cells["A4"].PutValue("Data3");
            cells["B4"].PutValue("Data4");

            // Configure JSON save options to skip empty rows
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // When SkipEmptyRows is true, empty rows are excluded from the output
                SkipEmptyRows = true
            };

            // Save the workbook as JSON using the configured options
            string outputPath = "output.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"JSON saved to '{outputPath}' with empty rows excluded.");
        }
    }
}
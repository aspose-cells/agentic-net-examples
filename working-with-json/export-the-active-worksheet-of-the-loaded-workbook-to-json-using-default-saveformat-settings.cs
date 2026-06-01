using System;
using Aspose.Cells;

namespace AsposeCellsJsonExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Determine the index of the active worksheet
            int activeSheetIndex = workbook.Worksheets.ActiveSheetIndex;

            // Create JSON save options with default settings
            JsonSaveOptions jsonOptions = new JsonSaveOptions();

            // Restrict export to the active worksheet only
            jsonOptions.SheetIndexes = new int[] { activeSheetIndex };

            // Save the active worksheet as a JSON file using default format settings
            string outputPath = "active_sheet.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Active worksheet exported to JSON at: {outputPath}");
        }
    }
}
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

            // Configure JSON save options to export only the active worksheet
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                // Export only the worksheet whose index matches the active sheet
                SheetIndexes = new int[] { activeSheetIndex }
            };

            // Save the active worksheet as a JSON file using default format settings
            string outputPath = "active_sheet.json";
            workbook.Save(outputPath, jsonOptions);

            Console.WriteLine($"Active worksheet exported to JSON at: {outputPath}");
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AddCellWatchesDemo
    {
        // Adds a cell to the watch window only if it is not already present.
        private static void AddWatchIfMissing(Worksheet sheet, string cellName)
        {
            // Check existing watches for the same cell name.
            foreach (CellWatch existingWatch in sheet.CellWatches)
            {
                if (string.Equals(existingWatch.CellName, cellName, StringComparison.OrdinalIgnoreCase))
                {
                    // Cell is already being watched; no need to add.
                    return;
                }
            }

            // Cell not found in the watch collection; add it.
            sheet.CellWatches.Add(cellName);
        }

        public static void Run()
        {
            // Path to the existing XLSX file.
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook from the file.
            Workbook workbook = new Workbook(inputPath);

            // Choose the worksheet to work with (first worksheet in this example).
            Worksheet sheet = workbook.Worksheets[0];

            // List of cell addresses that should be added to the watch window.
            string[] cellsToWatch = { "A1", "B2", "C3", "A1" }; // includes a duplicate for demonstration

            // Add each cell to the watch window, handling duplicates.
            foreach (string cellName in cellsToWatch)
            {
                AddWatchIfMissing(sheet, cellName);
            }

            // Optionally, display the total number of watches added.
            Console.WriteLine($"Total cell watches: {sheet.CellWatches.Count}");

            // Save the modified workbook.
            string outputPath = "OutputWorkbook_WithWatches.xlsx";
            workbook.Save(outputPath);
        }

        // Entry point for the application.
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
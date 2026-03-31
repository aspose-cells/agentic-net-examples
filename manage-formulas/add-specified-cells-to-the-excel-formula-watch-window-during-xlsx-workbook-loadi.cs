using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Define the cells that should be added to the watch window
            string[] cellsToWatch = { "B2", "C3", "D5" };

            // Iterate through each worksheet (or target a specific one)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Add each specified cell to the worksheet's watch collection
                foreach (string cellName in cellsToWatch)
                {
                    // The Add method returns the index of the newly added watch (optional)
                    int watchIndex = sheet.CellWatches.Add(cellName);
                    // Optionally, you can retrieve the CellWatch object for further inspection
                    CellWatch watch = sheet.CellWatches[watchIndex];
                    Console.WriteLine($"Added watch: Sheet='{sheet.Name}', Cell='{watch.CellName}' (Index={watchIndex})");
                }
            }

            // Save the workbook (the watch window information is persisted in the file)
            workbook.Save("output.xlsx");
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the workbook that contains the previously saved Watch Window configuration
                string watchConfigPath = "SavedWatchConfig.xlsx";

                // Load the workbook with the saved watch configuration if it exists
                Workbook watchConfigWorkbook;
                if (File.Exists(watchConfigPath))
                {
                    watchConfigWorkbook = new Workbook(watchConfigPath);
                }
                else
                {
                    Console.WriteLine($"Watch configuration file '{watchConfigPath}' not found. Continuing without applying watch config.");
                    watchConfigWorkbook = new Workbook(); // empty workbook
                }

                // Load or create the active workbook to which the watch configuration will be applied
                Workbook activeWorkbook = new Workbook();

                // Iterate through each worksheet in the saved configuration workbook
                foreach (Worksheet savedSheet in watchConfigWorkbook.Worksheets)
                {
                    // Ensure the active workbook has a corresponding worksheet
                    while (activeWorkbook.Worksheets.Count <= savedSheet.Index)
                    {
                        activeWorkbook.Worksheets.Add();
                    }

                    Worksheet activeSheet = activeWorkbook.Worksheets[savedSheet.Index];

                    // Copy each CellWatch from the saved sheet to the active sheet
                    foreach (CellWatch savedWatch in savedSheet.CellWatches)
                    {
                        // Add a new watch entry using the same cell name
                        int newWatchIndex = activeSheet.CellWatches.Add(savedWatch.CellName);

                        // Retrieve the newly added CellWatch object
                        CellWatch newWatch = activeSheet.CellWatches[newWatchIndex];

                        // Replicate the original watch properties
                        newWatch.Row = savedWatch.Row;
                        newWatch.Column = savedWatch.Column;
                        newWatch.CellName = savedWatch.CellName;
                    }
                }

                // Save the active workbook with the applied watch window configuration
                string outputPath = "ActiveWorkbook_WithWatchConfig.xlsx";
                activeWorkbook.Save(outputPath);

                Console.WriteLine($"Watch window configuration applied and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
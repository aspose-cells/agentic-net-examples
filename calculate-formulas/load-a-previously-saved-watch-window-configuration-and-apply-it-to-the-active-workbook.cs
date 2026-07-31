// Title: Apply a Saved Watch Window to an Active Workbook with Aspose.Cells (.NET)
// Description: The example checks for a watch‑configuration file and an active workbook, loads both using Aspose.Cells, clears existing CellWatches in each target sheet, copies the saved CellWatch entries from the source workbook, and saves the updated workbook with the new watch settings.
// Keywords: Aspose.Cells watch window | CellWatch copy .NET | load saved watch configuration | apply watch settings Excel | C# Aspose.Cells example | transfer CellWatches between workbooks
// Common Searches: How to copy a Watch Window from one Excel file to another using Aspose.Cells | Aspose.Cells .NET code to transfer CellWatch objects between workbooks | Load and apply a saved watch configuration in C# with Aspose.Cells | Copy watch list across multiple generated reports Aspose.Cells
// Developer Intent: Load a previously saved Watch Window and apply its CellWatch list to the currently opened workbook.
// Use Cases: Standardize watch lists across a suite of financial models before distribution. | Retain monitoring of key cells when opening a new version of a spreadsheet. | Automate the injection of a predefined watch configuration into batch‑generated reports.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy CellWatch entries from a source workbook to a target workbook, handling missing files and worksheet mismatches. | Create robust error handling for differing worksheet counts when applying a saved Watch Window with Aspose.Cells. | Explain step‑by‑step how CellWatches are cleared and re‑added in Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    // The example checks for a watch‑configuration file and an active workbook, loads both using Aspose.Cells, clears existing CellWatches in each target sheet, copies the saved CellWatch entries from the source workbook, and saves the updated workbook with the new watch settings.
    class Program
    {
        static void Main()
        {
            try
            {
                const string watchConfigPath = "WatchConfig.xlsx";
                const string activeWorkbookPath = "ActiveWorkbook.xlsx";
                const string outputPath = "ActiveWorkbook_WithWatch.xlsx";

                // Verify that the required input files exist
                if (!File.Exists(watchConfigPath))
                {
                    Console.WriteLine($"Error: File '{watchConfigPath}' not found.");
                    return;
                }

                if (!File.Exists(activeWorkbookPath))
                {
                    Console.WriteLine($"Error: File '{activeWorkbookPath}' not found.");
                    return;
                }

                // Load the workbook that contains the previously saved Watch Window configuration
                Workbook watchConfigWorkbook = new Workbook(watchConfigPath);

                // Load the active workbook to which the watch configuration will be applied
                Workbook activeWorkbook = new Workbook(activeWorkbookPath);

                // Iterate through each worksheet (assuming both workbooks have the same sheet order)
                for (int i = 0; i < watchConfigWorkbook.Worksheets.Count; i++)
                {
                    Worksheet sourceSheet = watchConfigWorkbook.Worksheets[i];
                    Worksheet targetSheet = activeWorkbook.Worksheets[i];

                    // Clear any existing watches in the target sheet
                    targetSheet.CellWatches.Clear();

                    // Copy each CellWatch from the source sheet to the target sheet
                    foreach (CellWatch sourceWatch in sourceSheet.CellWatches)
                    {
                        // Add a new watch in the target sheet using the same cell name
                        targetSheet.CellWatches.Add(sourceWatch.CellName);
                    }
                }

                // Save the active workbook with the applied Watch Window configuration
                activeWorkbook.Save(outputPath);
                Console.WriteLine($"Watch configuration applied and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}

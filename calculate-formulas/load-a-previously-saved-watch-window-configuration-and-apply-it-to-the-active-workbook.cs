// Title: Load a saved watch window configuration from a text file and apply it to every worksheet in an active workbook with Aspose.Cells for .NET
// AI Prompts: Read cell addresses line‑by‑line from a .txt file and add each address as a CellWatch to all worksheets in a loaded workbook using Aspose.Cells. | Validate the existence of the workbook and the watch‑list file, then handle any missing‑file errors before applying watches. | After adding the CellWatch entries, save the workbook to a new file that includes the applied watch configuration.
// Common Searches: C# Aspose.Cells load watch window configuration from txt file | How to add multiple CellWatch entries to every worksheet programmatically | Saving a workbook after applying a watch list with Aspose.Cells .NET | Read watch list from text file and set CellWatches in Aspose.Cells
// Tags: load watch configuration Aspose.Cells | add cell watches from text file | apply watch list to all worksheets | save workbook with cell watches | cellwatches collection usage .NET

using System;
using System.IO;
using Aspose.Cells;

namespace WatchWindowDemo
{
    // The example loads an existing workbook, reads cell addresses from a text file, adds each address as a CellWatch to every worksheet via Aspose.Cells, and saves the workbook with the new watch configuration.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that is currently active (to which we want to apply the watch configuration)
            string workbookPath = "ActiveWorkbook.xlsx";

            // Path to the saved watch window configuration file.
            // The file is expected to contain one cell address per line, e.g.:
            // B2
            // C5
            // D10
            string configPath = "WatchConfig.txt";

            try
            {
                // Verify that the workbook file exists before attempting to load it.
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the active workbook.
                Workbook workbook = new Workbook(workbookPath);

                // Verify that the configuration file exists.
                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"Configuration file not found: {configPath}");
                    return;
                }

                // Read the watch configuration (one cell address per line).
                string[] watchLines = File.ReadAllLines(configPath);

                // Apply the watch configuration to each worksheet in the workbook.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (string line in watchLines)
                    {
                        string cellAddress = line.Trim();

                        // Skip empty lines.
                        if (string.IsNullOrEmpty(cellAddress))
                            continue;

                        // Add the cell to the watch window.
                        // The Add method returns the index of the newly added watch item.
                        int watchIndex = sheet.CellWatches.Add(cellAddress);

                        // Retrieve the CellWatch object (optional verification/modification).
                        CellWatch watch = sheet.CellWatches[watchIndex];
                        watch.CellName = cellAddress; // Ensure the name matches the address.
                    }
                }

                // Save the workbook with the applied watch configuration.
                string outputPath = "ActiveWorkbook_WithWatches.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Watch configuration applied and workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

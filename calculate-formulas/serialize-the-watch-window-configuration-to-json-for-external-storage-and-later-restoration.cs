// Title: C# – Serialize and Restore Aspose.Cells Watch Window to JSON
// Description: Demonstrates how to capture the CellWatch items of a worksheet, serialize them to an indented JSON file with System.Text.Json, and later deserialize the file to re‑add the watches to a new workbook. Includes saving and loading the configuration and optional workbook export.
// Keywords: Aspose.Cells | CellWatch | watch window | JSON serialization | C# example | System.Text.Json | export watch configuration | import watch configuration | persist watch items | restore watch window
// Common Searches: Aspose.Cells serialize watch window to JSON | how to save CellWatch list as JSON in C# | restore Aspose.Cells watch items from file | export and import watch window configuration | C# example for persisting CellWatch objects
// Developer Intent: Export the worksheet's watch window to a JSON file and reload it later into another workbook.
// Use Cases: Preserve user‑defined watch items between application sessions. | Share a predefined watch configuration with teammates. | Create a backup of watch settings before running bulk calculations.
// AI Prompts: Write C# code that reads a JSON file of CellWatch objects and adds them to a worksheet using Aspose.Cells. | Show how to customize JsonSerializerOptions for Aspose.Cells CellWatch serialization. | Explain a method to clear existing watches in a worksheet before re‑importing them from JSON.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsWatchWindowJson
{
    // Demonstrates how to capture the CellWatch items of a worksheet, serialize them to an indented JSON file with System.Text.Json, and later deserialize the file to re‑add the watches to a new workbook. Includes saving and loading the configuration and optional workbook export.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a couple of cell watch items to the watch window
            int watchIndex1 = sheet.CellWatches.Add("B2");
            CellWatch watch1 = sheet.CellWatches[watchIndex1];
            watch1.Row = 1;          // zero‑based row index
            watch1.Column = 1;       // zero‑based column index
            watch1.CellName = "B2";

            int watchIndex2 = sheet.CellWatches.Add("D5");
            CellWatch watch2 = sheet.CellWatches[watchIndex2];
            watch2.Row = 4;
            watch2.Column = 3;
            watch2.CellName = "D5";

            // ------------------------------------------------------------
            // Serialize the current watch window configuration to JSON
            // ------------------------------------------------------------
            var watchList = new List<CellWatch>();
            for (int i = 0; i < sheet.CellWatches.Count; i++)
            {
                watchList.Add(sheet.CellWatches[i]);
            }

            string json = JsonSerializer.Serialize(
                watchList,
                new JsonSerializerOptions { WriteIndented = true });

            // Save the JSON to an external file
            string jsonPath = "WatchWindowConfig.json";
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Watch window configuration saved to: {jsonPath}");
            Console.WriteLine(json);

            // ------------------------------------------------------------
            // Demonstrate restoration of the watch window from the saved JSON
            // ------------------------------------------------------------
            // (In a real scenario you might load a different workbook)
            // Clear existing watches for demonstration purposes
            // Note: Aspose.Cells does not provide a direct Clear method,
            // so we recreate the worksheet to start fresh.
            workbook = new Workbook();               // new workbook
            sheet = workbook.Worksheets[0];          // first worksheet

            // Load the JSON file
            string loadedJson = File.ReadAllText(jsonPath);
            List<CellWatch> loadedWatches = JsonSerializer.Deserialize<List<CellWatch>>(loadedJson);

            // Re‑add each watch item to the worksheet's watch window
            foreach (CellWatch cw in loadedWatches)
            {
                // Adding by cell name automatically sets row/column internally
                sheet.CellWatches.Add(cw.CellName);
            }

            // Save the workbook to verify that watches are restored (optional)
            workbook.Save("RestoredWorkbook.xlsx");
            Console.WriteLine("Workbook saved with restored watch window.");
        }
    }
}

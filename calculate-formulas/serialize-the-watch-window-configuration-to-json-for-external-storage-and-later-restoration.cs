// Title: How to serialize Aspose.Cells watch window configuration to JSON and restore it in C#
// AI Prompts: Write a C# method that loops through Worksheet.CellWatches, creates a DTO with CellName, Row, and Column, and saves the collection as an indented JSON file using System.Text.Json. | Develop a C# routine that reads a JSON file containing watch entries, clears the existing CellWatches on a worksheet, re‑adds each watch (using the stored cell name or converting row/column to an address), and saves the workbook.
// Common Searches: Aspose.Cells export watch window to JSON file C# | C# restore cell watches from JSON using Aspose.Cells | Save and load worksheet watch list Aspose.Cells .NET | How to persist Aspose.Cells CellWatches across sessions | Serialize Aspose.Cells watch configuration for version control
// Tags: Aspose.Cells serialize watch window to JSON | C# deserialize CellWatches from JSON | Aspose.Cells watch list persistence | JSON export of worksheet CellWatches | Reapply Aspose.Cells watch configuration programmatically

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsWatchWindowJsonDemo
{
    // Simple DTO for serializing CellWatch information
    // This example shows how to capture the CellWatches of the first worksheet in an Aspose.Cells workbook, serialize each watch's name, row, and column to a formatted JSON file, and later deserialize that JSON to rebuild the watch list and save the workbook.
    public class CellWatchInfo
    {
        public string CellName { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
    }

    public class WatchWindowJsonHandler
    {
        // Serialize the current watch window configuration to a JSON file
        public static void SaveWatchWindowConfig(string workbookPath, string jsonPath)
        {
            // Load or create a workbook
            Workbook workbook = File.Exists(workbookPath) ? new Workbook(workbookPath) : new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: add some watch items if none exist
            if (sheet.CellWatches.Count == 0)
            {
                sheet.CellWatches.Add("B2");
                sheet.CellWatches.Add("C5");
            }

            // Collect watch information into a list of DTOs
            List<CellWatchInfo> watchList = new List<CellWatchInfo>();
            for (int i = 0; i < sheet.CellWatches.Count; i++)
            {
                CellWatch cw = sheet.CellWatches[i];
                watchList.Add(new CellWatchInfo
                {
                    CellName = cw.CellName,
                    Row = cw.Row,
                    Column = cw.Column
                });
            }

            // Serialize with indentation for readability
            JsonSerializerOptions jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string json = JsonSerializer.Serialize(watchList, jsonOptions);

            // Write JSON to the specified file
            File.WriteAllText(jsonPath, json);
        }

        // Restore the watch window configuration from a JSON file
        public static void RestoreWatchWindowConfig(string workbookPath, string jsonPath)
        {
            // Load the workbook (must exist)
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Read and deserialize the JSON file
            string json = File.ReadAllText(jsonPath);
            List<CellWatchInfo> watchList = JsonSerializer.Deserialize<List<CellWatchInfo>>(json);

            // Clear existing watches (optional, depending on desired behavior)
            sheet.CellWatches.Clear();

            // Re‑add each watch item
            foreach (CellWatchInfo info in watchList)
            {
                // Add by cell name; if CellName is null/empty, construct from row/column
                if (!string.IsNullOrEmpty(info.CellName))
                {
                    sheet.CellWatches.Add(info.CellName);
                }
                else
                {
                    // Convert zero‑based row/column to Excel style address (e.g., B2)
                    string address = CellsHelper.CellIndexToName(info.Row, info.Column);
                    sheet.CellWatches.Add(address);
                }
            }

            // Save the workbook after restoration
            workbook.Save(workbookPath);
        }

        // Demonstration entry point
        public static void RunDemo()
        {
            string workbookFile = "WatchDemo.xlsx";
            string jsonFile = "WatchConfig.json";

            // Create a workbook and add some data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample");
            ws.Cells["B2"].PutValue(123);
            ws.Cells["C5"].PutValue("WatchMe");
            wb.Save(workbookFile);

            // Save watch configuration to JSON
            SaveWatchWindowConfig(workbookFile, jsonFile);
            Console.WriteLine($"Watch configuration saved to {jsonFile}");

            // Modify workbook (remove watches) to demonstrate restoration
            Workbook wb2 = new Workbook(workbookFile);
            wb2.Worksheets[0].CellWatches.Clear();
            wb2.Save(workbookFile);

            // Restore watch configuration from JSON
            RestoreWatchWindowConfig(workbookFile, jsonFile);
            Console.WriteLine("Watch configuration restored from JSON.");
        }
    }

    class Program
    {
        static void Main()
        {
            WatchWindowJsonHandler.RunDemo();
        }
    }
}

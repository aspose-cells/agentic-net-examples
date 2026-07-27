using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsWatchWindowJsonDemo
{
    // Simple DTO representing a watch item for JSON serialization
    public class WatchItem
    {
        public string CellName { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and add some cell watches to the first sheet
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add watches for cells B2 and C3
            sheet.CellWatches.Add("B2");
            sheet.CellWatches.Add("C3");

            // -----------------------------------------------------------------
            // 2. Extract the watch configuration into a serializable list
            // -----------------------------------------------------------------
            List<WatchItem> watchItems = new List<WatchItem>();
            for (int i = 0; i < sheet.CellWatches.Count; i++)
            {
                CellWatch cw = sheet.CellWatches[i];
                watchItems.Add(new WatchItem
                {
                    CellName = cw.CellName,
                    Row = cw.Row,
                    Column = cw.Column
                });
            }

            // -----------------------------------------------------------------
            // 3. Serialize the watch list to JSON and save it to a file
            // -----------------------------------------------------------------
            JsonSerializerOptions jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            string json = JsonSerializer.Serialize(watchItems, jsonOptions);
            string jsonPath = "WatchWindowConfig.json";
            File.WriteAllText(jsonPath, json);
            Console.WriteLine($"Watch window configuration saved to '{jsonPath}':");
            Console.WriteLine(json);

            // -----------------------------------------------------------------
            // 4. Demonstrate restoration: create a new workbook and load watches
            // -----------------------------------------------------------------
            Workbook restoredWorkbook = new Workbook();
            Worksheet restoredSheet = restoredWorkbook.Worksheets[0];

            // Load JSON from file
            string loadedJson = File.ReadAllText(jsonPath);
            List<WatchItem> loadedWatches = JsonSerializer.Deserialize<List<WatchItem>>(loadedJson);

            // Re‑add each watch to the worksheet
            foreach (WatchItem item in loadedWatches)
            {
                // The Add method accepts a cell name; we can use the stored CellName
                restoredSheet.CellWatches.Add(item.CellName);
            }

            // Save the restored workbook to verify that watches are present
            string restoredPath = "RestoredWorkbook.xlsx";
            restoredWorkbook.Save(restoredPath);
            Console.WriteLine($"Restored workbook saved to '{restoredPath}'.");
        }
    }
}
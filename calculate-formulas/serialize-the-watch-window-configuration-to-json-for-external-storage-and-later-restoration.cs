using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class WatchWindowJsonDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["C3"].PutValue(30);

        // Add cells to the watch window
        sheet.CellWatches.Add("A1");
        sheet.CellWatches.Add("B2");
        sheet.CellWatches.Add("C3");

        // Gather all CellWatch objects into a list
        List<CellWatch> watchList = new List<CellWatch>();
        for (int i = 0; i < sheet.CellWatches.Count; i++)
        {
            watchList.Add(sheet.CellWatches[i]);
        }

        // Serialize the watch configuration to JSON
        string json = JsonSerializer.Serialize(watchList, new JsonSerializerOptions { WriteIndented = true });
        string jsonPath = "watchConfig.json";
        File.WriteAllText(jsonPath, json);
        Console.WriteLine($"Watch configuration saved to {jsonPath}");
        Console.WriteLine(json);

        // -------------------- Restoration --------------------
        // Load the JSON back
        string loadedJson = File.ReadAllText(jsonPath);
        List<CellWatch> loadedWatches = JsonSerializer.Deserialize<List<CellWatch>>(loadedJson);

        // Optionally clear existing watches
        sheet.CellWatches.Clear();

        // Re‑add watches from the deserialized list
        foreach (CellWatch cw in loadedWatches)
        {
            // Prefer the stored CellName; if missing, construct it from row/column
            string cellName = !string.IsNullOrEmpty(cw.CellName)
                ? cw.CellName
                : CellsHelper.CellIndexToName(cw.Row, cw.Column);
            sheet.CellWatches.Add(cellName);
        }

        // Save the workbook to verify that watches are restored
        workbook.Save("RestoredWatchWindow.xlsx");
        Console.WriteLine("Workbook saved with restored watch window.");
    }
}
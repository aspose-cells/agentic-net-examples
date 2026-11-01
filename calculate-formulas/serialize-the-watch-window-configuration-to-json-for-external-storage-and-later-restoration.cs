// Title: Serialize and restore Aspose.Cells Watch Window configuration to JSON in C#
// Common Searches: C# Aspose.Cells export WatchWindow items to a JSON file | deserialize WatchWindow configuration from JSON using Aspose.Cells .NET | save Excel watch window settings with Aspose.Cells and reload later | example of Aspose.Cells WatchWindow serialization in C# | use reflection to access the WatchWindow property in an Aspose.Cells worksheet | persist worksheet watch items to JSON and restore with Aspose.Cells
// Tags: Aspose.Cells WatchWindow JSON serialization | C# serialize worksheet watch items | Aspose.Cells restore watch window from JSON | reflection access WatchWindow property Aspose.Cells | persist Excel watch window configuration .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The WatchWindowSerializer class demonstrates how to use reflection to obtain a worksheet's WatchWindow, convert its WatchItems collection into a simple DTO list, serialize the list to an indented JSON file, and later deserialize the JSON to recreate the watch items, enabling persistent storage and restoration of Excel watch window settings with Aspose.Cells for .NET.
class WatchWindowSerializer
{
    // Serialize the Watch Window configuration of the first worksheet to a JSON file.
    public static void SerializeWatchWindow(string workbookPath, string jsonPath)
    {
        try
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Use reflection to obtain the WatchWindow (may not exist in older versions).
            var watchWindowProp = sheet.GetType().GetProperty("WatchWindow");
            if (watchWindowProp == null)
                throw new NotSupportedException("WatchWindow feature is not available in this Aspose.Cells version.");

            dynamic watchWindow = watchWindowProp.GetValue(sheet);
            if (watchWindow == null)
                throw new InvalidOperationException("Unable to retrieve WatchWindow.");

            // Retrieve the collection of watch items.
            var watchItems = watchWindow.WatchItems as IEnumerable<object>;
            if (watchItems == null)
                throw new InvalidOperationException("WatchItems collection is unavailable.");

            // Convert to DTO list for JSON serialization.
            List<WatchItemDto> items = new List<WatchItemDto>();
            foreach (dynamic watchItem in watchItems)
            {
                items.Add(new WatchItemDto
                {
                    CellName = watchItem.CellName,
                    Comment = watchItem.Comment
                });
            }

            // Serialize with indentation.
            string json = JsonSerializer.Serialize(items, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during serialization: {ex.Message}");
        }
    }

    // Restore a Watch Window configuration from a JSON file into the first worksheet of a workbook.
    public static void RestoreWatchWindow(string workbookPath, string jsonPath, string outputWorkbookPath)
    {
        try
        {
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook not found: {workbookPath}");
            if (!File.Exists(jsonPath))
                throw new FileNotFoundException($"JSON file not found: {jsonPath}");

            // Load the workbook.
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Obtain WatchWindow via reflection.
            var watchWindowProp = sheet.GetType().GetProperty("WatchWindow");
            if (watchWindowProp == null)
                throw new NotSupportedException("WatchWindow feature is not available in this Aspose.Cells version.");

            dynamic watchWindow = watchWindowProp.GetValue(sheet);
            if (watchWindow == null)
                throw new InvalidOperationException("Unable to retrieve WatchWindow.");

            // Clear existing items.
            watchWindow.Clear();

            // Read and deserialize JSON.
            string json = File.ReadAllText(jsonPath);
            List<WatchItemDto>? items = JsonSerializer.Deserialize<List<WatchItemDto>>(json);
            if (items == null)
                throw new InvalidOperationException("Deserialized watch items are null.");

            // Re‑add each watch item.
            foreach (WatchItemDto dto in items)
            {
                watchWindow.Add(dto.CellName, dto.Comment);
            }

            // Save the modified workbook.
            workbook.Save(outputWorkbookPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during restoration: {ex.Message}");
        }
    }

    // Simple DTO used for JSON serialization.
    private class WatchItemDto
    {
        public string CellName { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
    }
}

// Example usage.
class Program
{
    static void Main()
    {
        string sourceWorkbook = "Source.xlsx";
        string jsonFile = "WatchWindowConfig.json";
        string restoredWorkbook = "Restored.xlsx";

        // Serialize current Watch Window to JSON.
        WatchWindowSerializer.SerializeWatchWindow(sourceWorkbook, jsonFile);

        // Restore Watch Window from JSON into a new workbook copy.
        WatchWindowSerializer.RestoreWatchWindow(sourceWorkbook, jsonFile, restoredWorkbook);
    }
}

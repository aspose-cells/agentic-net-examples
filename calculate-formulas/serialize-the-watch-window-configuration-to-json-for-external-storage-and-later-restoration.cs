// Title: Serialize Aspose.Cells Watch Window Data to JSON and Restore It in C# (.NET)
// Common Searches: how to save Aspose.Cells watch window items to a JSON file in C# | C# code to export Excel cell addresses, values and formulas to JSON using Aspose.Cells | restore watch window configuration from JSON with Aspose.Cells for .NET | example of reading JSON to recreate watched cells in an Excel workbook with Aspose.Cells | Aspose.Cells serialize cell formulas and values to JSON for external storage
// Tags: Aspose.Cells watch window JSON serialization .NET | C# export Excel cell values and formulas to JSON | restore Excel watch items from JSON Aspose.Cells | JSON storage for Aspose.Cells watch configuration | serialize cell address value formula list C#

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace WatchWindowSerialization
{
    // DTO for JSON representation of a watch item
    // The sample creates or loads an Excel workbook, gathers each populated cell's address, value, and formula into a DTO list, serializes the list to an indented JSON file, then reads the JSON back, reconstructs the cells' values or formulas, and saves the restored workbook using Aspose.Cells for .NET.
    public class WatchItemDto
    {
        public string? Name { get; set; }      // Cell address (e.g., "A1")
        public string? Value { get; set; }     // String representation of the cell value
        public string? Formula { get; set; }   // Cell formula, if any
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string jsonPath = "watchWindow.json";
            const string outputPath = "output_restored.xlsx";

            try
            {
                // Ensure the input workbook exists; create a simple one if missing
                if (!File.Exists(inputPath))
                {
                    var wb = new Workbook();
                    var ws = wb.Worksheets[0];
                    ws.Cells["A1"].PutValue(123);
                    ws.Cells["B2"].Formula = "=A1*2";
                    wb.Save(inputPath);
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);
                var sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Serialize Watch Window configuration to JSON file
                // -------------------------------------------------
                var watchItems = new List<WatchItemDto>();

                // Since Aspose.Cells for .NET does not expose a WatchWindow API,
                // we treat all cells with a value or formula as "watched" items.
                var cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var cell = cells[row, col];
                        if (cell != null &&
                            (cell.Type != CellValueType.IsNull || !string.IsNullOrEmpty(cell.Formula)))
                        {
                            watchItems.Add(new WatchItemDto
                            {
                                Name = cell.Name,
                                Value = cell.Value?.ToString(),
                                Formula = cell.Formula
                            });
                        }
                    }
                }

                // Convert the list to JSON
                string json = JsonSerializer.Serialize(watchItems, new JsonSerializerOptions { WriteIndented = true });

                // Save JSON to external storage
                File.WriteAllText(jsonPath, json);

                // -------------------------------------------------
                // Restore Watch Window configuration from JSON file
                // -------------------------------------------------
                if (!File.Exists(jsonPath))
                    throw new FileNotFoundException($"JSON file not found: {jsonPath}");

                string loadedJson = File.ReadAllText(jsonPath);
                var loadedItems = JsonSerializer.Deserialize<List<WatchItemDto>>(loadedJson) ?? new List<WatchItemDto>();

                // Recreate cells based on the stored data
                foreach (var dto in loadedItems)
                {
                    if (string.IsNullOrEmpty(dto.Name))
                        continue; // Skip invalid entries

                    var cell = sheet.Cells[dto.Name];
                    if (!string.IsNullOrEmpty(dto.Formula))
                    {
                        cell.Formula = dto.Formula;
                    }
                    else if (!string.IsNullOrEmpty(dto.Value))
                    {
                        // Attempt to parse numeric values; otherwise store as string
                        if (double.TryParse(dto.Value, out double numeric))
                            cell.PutValue(numeric);
                        else
                            cell.PutValue(dto.Value);
                    }
                }

                // Save the workbook after restoration
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook restored and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

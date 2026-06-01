using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

namespace AsposeCellsJsonBackup
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can iterate over all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare a list to hold address‑value pairs
            var cellData = new List<Dictionary<string, string>>();

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Enumerate each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    // Skip cells that are completely empty
                    if (cell == null || cell.Type == CellValueType.IsNull)
                        continue;

                    // Create a simple key‑value representation
                    var entry = new Dictionary<string, string>
                    {
                        { "Address", cell.Name },               // e.g., "A1"
                        { "Value", cell.Value?.ToString() ?? "" } // cell value as string
                    };
                    cellData.Add(entry);
                }
            }

            // Serialize the list to JSON with indentation for readability
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(cellData, jsonOptions);

            // Save the JSON to a file (backup)
            string outputPath = "worksheet_backup.json";
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"Worksheet data has been backed up to '{outputPath}'.");
        }
    }
}
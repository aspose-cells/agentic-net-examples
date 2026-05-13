using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file
        string inputPath = "input.xlsb";

        // Path where the JSON representation will be saved
        string jsonOutputPath = "output.json";

        // Load the XLSB workbook using LoadOptions (lifecycle rule)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Calculate all formulas in the workbook (lifecycle rule)
        workbook.CalculateFormula();

        // Prepare a serializable structure for the workbook data
        var workbookData = new List<Dictionary<string, object>>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var sheetInfo = new Dictionary<string, object>
            {
                ["Name"] = sheet.Name
            };

            var rows = new List<List<object>>();

            // Determine the used range of the worksheet
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int r = 0; r <= maxRow; r++)
            {
                var row = new List<object>();
                for (int c = 0; c <= maxCol; c++)
                {
                    var cell = sheet.Cells[r, c];
                    // Use the calculated value of the cell; null if empty
                    row.Add(cell.Value);
                }
                rows.Add(row);
            }

            sheetInfo["Data"] = rows;
            workbookData.Add(sheetInfo);
        }

        // Serialize the structure to formatted JSON
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(workbookData, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(jsonOutputPath, json);

        Console.WriteLine($"JSON representation saved to: {jsonOutputPath}");
    }
}
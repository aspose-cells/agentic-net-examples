using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Paths for input XLSB file and output JSON file
        string inputPath = "input.xlsb";
        string outputPath = "output.json";

        // Load the XLSB workbook with appropriate load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsb);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Prepare a structure to hold the workbook data
        var workbookData = new Dictionary<string, object>();

        // Iterate through each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var rows = new List<List<object>>();

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Extract cell values row by row
            for (int r = 0; r <= maxRow; r++)
            {
                var rowData = new List<object>();
                for (int c = 0; c <= maxCol; c++)
                {
                    var cell = sheet.Cells[r, c];
                    // Use the cell's value; if null, store null
                    rowData.Add(cell.Value);
                }
                rows.Add(rowData);
            }

            // Add the sheet data to the overall dictionary
            workbookData[sheet.Name] = rows;
        }

        // Serialize the data to JSON with indentation for readability
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(workbookData, jsonOptions);

        // Write the JSON string to the output file
        File.WriteAllText(outputPath, json);

        Console.WriteLine($"Workbook data has been exported to JSON file: {outputPath}");
    }
}
// Title: Load an XLSB workbook, recalculate all formulas, and export its contents to JSON with Aspose.Cells for .NET
// Common Searches: how to load an XLSB file with Aspose.Cells and recalculate formulas in C# | export Aspose.Cells worksheet values to a JSON file using .NET | C# code to read all cells from an XLSB workbook and convert to JSON | Aspose.Cells calculate all formulas before saving workbook data as JSON
// Tags: XLSB workbook loading Aspose.Cells .NET | formula recalculation Aspose.Cells | worksheet data serialization to JSON C# | cell iteration Aspose.Cells workbook | Aspose.Cells export workbook to JSON format

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// The example loads an XLSB workbook with Aspose.Cells using LoadOptions, forces a full formula recalculation, iterates through each worksheet and cell to collect values, serializes the collected data into an indented JSON structure with System.Text.Json, ensures the output directory exists, and writes the JSON to a file while handling possible errors.
class Program
{
    static void Main(string[] args)
    {
        // Paths for input XLSB and output JSON
        string inputPath = "input.xlsb";
        string outputPath = "output.json";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the XLSB workbook with appropriate load options
            var loadOptions = new LoadOptions(LoadFormat.Xlsb);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate all formulas to ensure up‑to‑date values
            workbook.CalculateFormula();

            // Collect workbook data for JSON serialization
            var workbookData = new List<WorksheetData>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var sheetData = new WorksheetData
                {
                    Name = sheet.Name,
                    Rows = new List<List<object>>()
                };

                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    var rowData = new List<object>();
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        rowData.Add(cell.Value);
                    }
                    sheetData.Rows.Add(rowData);
                }

                workbookData.Add(sheetData);
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Serialize to indented JSON using System.Text.Json
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(workbookData, jsonOptions);

            // Write JSON to file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"JSON representation saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// DTO to hold worksheet information
public class WorksheetData
{
    public string Name { get; set; } = string.Empty;
    public List<List<object>> Rows { get; set; } = new List<List<object>>();
}

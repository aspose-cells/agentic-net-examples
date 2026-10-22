// Title: Convert an XLSB workbook to formatted JSON after recalculating formulas with Aspose.Cells for .NET
// AI Prompts: Load an XLSB file using Aspose.Cells, call Workbook.CalculateFormula, iterate through each worksheet's used range, and serialize the sheet name, row index, cell address, and evaluated value to a pretty‑printed JSON file. | Write a C# program that opens an XLSB workbook, forces formula evaluation, extracts all cell values (including null for empty cells) with their addresses, and saves the complete workbook structure as indented JSON.
// Common Searches: aspnet convert xlsb workbook to json with formula evaluation using Aspose.Cells | c# recalculate all formulas in an XLSB file before exporting data to JSON | how to export Aspose.Cells worksheet data with cell addresses to a formatted JSON file | serialize entire workbook including empty cells to JSON in .NET
// Tags: XLSB workbook formula recalculation Aspose.Cells | export worksheet cells to JSON C# | Aspose.Cells load XLSB with LoadOptions | serialize workbook data as indented JSON | cell address value extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using System.Text.Json;

// The program loads an XLSB workbook via Aspose.Cells, forces full formula calculation, walks through each worksheet's used range to capture cell addresses and their evaluated values (null for empty cells), builds a hierarchical object containing sheet names and rows, and writes the result as a pretty‑printed JSON file.
class Program
{
    static void Main(string[] args)
    {
        // Paths for input XLSB and output JSON
        string inputPath = "input.xlsb";
        string outputPath = "output.json";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
            return;
        }

        try
        {
            // Load the XLSB workbook with appropriate load options
            var loadOptions = new LoadOptions(LoadFormat.Xlsb);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Container for the entire workbook data
            var workbookData = new List<object>();

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var sheetData = new Dictionary<string, object>
                {
                    ["Name"] = sheet.Name
                };

                var rowsData = new List<Dictionary<string, object>>();

                // Determine the used range of the sheet
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // Iterate through each row within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    var rowData = new Dictionary<string, object>();

                    // Iterate through each column within the used range
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        Cell cell = cells[row, col];
                        string address = cell.Name; // e.g., "A1"

                        // Use Cell.Type to determine emptiness (CellValueType.IsNull indicates an empty cell)
                        object value = cell.Type == CellValueType.IsNull ? null : cell.Value;
                        rowData[address] = value;
                    }

                    rowsData.Add(rowData);
                }

                sheetData["Rows"] = rowsData;
                workbookData.Add(sheetData);
            }

            // Serialize the workbook data to JSON with indentation
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(workbookData, jsonOptions);

            // Write the JSON string to the output file
            File.WriteAllText(outputPath, json);

            Console.WriteLine($"JSON representation saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

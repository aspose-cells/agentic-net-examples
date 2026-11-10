// Title: Export all worksheet formulas to a JSON file using Aspose.Cells for .NET (C#)
// Common Searches: how to extract every formula from an Excel workbook and save as JSON with Aspose.Cells C# | C# code to list cell formulas across all sheets and write to a JSON file | Aspose.Cells iterate MaxDataRow MaxDataColumn to collect formulas for external analysis | serialize Excel formula collection to indented JSON using System.Text.Json in .NET | command‑line tool to export Excel formulas to JSON using Aspose.Cells
// Tags: Aspose.Cells export formulas to JSON | C# iterate workbook cells with MaxDataRow | collect non‑empty cell formulas across worksheets | System.Text.Json serialize formula list | write JSON output file in .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

// The example loads a workbook with Aspose.Cells, walks through each worksheet and every cell up to the maximum data row and column, captures cells that contain a formula together with their sheet name and address, creates a list of these objects, ensures the destination folder exists, serializes the list to indented JSON using System.Text.Json, and writes the JSON to a specified file. Command‑line arguments allow custom input and output paths, and errors are reported to the console.
class Program
{
    static void Main(string[] args)
    {
        // Default placeholders; can be overridden by command‑line arguments
        string workbookPath = "{WorkbookPath}";
        string jsonPath = "{JsonPath}";

        if (args.Length >= 1) workbookPath = args[0];
        if (args.Length >= 2) jsonPath = args[1];

        try
        {
            // Verify the workbook file exists before loading
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException($"Workbook file not found: {workbookPath}");

            var workbook = new Workbook(workbookPath);
            var formulaList = new List<object>();

            // Iterate through all worksheets and collect formulas
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        var cell = cells[row, col];
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            formulaList.Add(new
                            {
                                SheetName = sheet.Name,
                                Address = cell.Name,
                                Formula = cell.Formula
                            });
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string jsonDir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(jsonDir) && !Directory.Exists(jsonDir))
                Directory.CreateDirectory(jsonDir);

            // Serialize formulas to JSON and write to file
            var json = JsonSerializer.Serialize(formulaList, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonPath, json);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

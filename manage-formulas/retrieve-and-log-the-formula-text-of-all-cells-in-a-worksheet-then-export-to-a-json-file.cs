// Title: Extract every cell’s formula from an Excel worksheet and save the results as a JSON file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get formula text for each cell in a worksheet | How to save all Excel formulas into a JSON file with Aspose.Cells | Iterate over used range and retrieve formulas using Aspose.Cells .NET | Save Excel cell formulas as a JSON dictionary in C# | Log cell addresses and their formulas while exporting to JSON in C#
// Tags: Aspose.Cells retrieve cell formulas | C# generate JSON file from formulas | Aspose.Cells iterate used range | Serialize dictionary to JSON in C# | Log cell address and formula Aspose.Cells | Convert Excel formulas to JSON format

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, selects the first worksheet, determines its used range, and iterates through every cell. For each cell it captures the formula text (or an empty string if none), logs the cell address and formula, and stores the pair in a dictionary. The dictionary is then serialized to an indented JSON string and written to a specified file, with checks for file existence and output directory creation, plus comprehensive error handling.
class Program
{
    static void Main()
    {
        // Paths for input Excel and output JSON
        string excelPath = @"C:\Data\SourceWorkbook.xlsx";
        string jsonPath = @"C:\Data\CellFormulas.json";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The source Excel file was not found at '{excelPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Select the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Dictionary to hold cell address and its formula text
            Dictionary<string, string> cellFormulas = new Dictionary<string, string>();

            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxColumn = worksheet.Cells.MaxDataColumn;

            // Iterate through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Retrieve the formula text (empty string if the cell has no formula)
                    string formula = cell.IsFormula ? cell.Formula : string.Empty;

                    // Log the cell address and its formula
                    Console.WriteLine($"{cell.Name}: {formula}");

                    // Store the formula in the dictionary (include all cells, even those without formulas)
                    cellFormulas[cell.Name] = formula;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(jsonPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Serialize the dictionary to JSON with indentation
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonContent = JsonSerializer.Serialize(cellFormulas, jsonOptions);

            // Write the JSON content to the output file
            File.WriteAllText(jsonPath, jsonContent);

            Console.WriteLine($"Formula data exported to JSON file at: {jsonPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

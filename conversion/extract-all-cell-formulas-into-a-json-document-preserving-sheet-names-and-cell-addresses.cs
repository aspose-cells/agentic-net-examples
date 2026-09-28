// Title: Extract all formulas from an Excel workbook to a JSON file with sheet names and cell addresses using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, scans every worksheet, collects each cell's address and its Formula property, and saves the collection as a pretty‑printed JSON document. | Modify the program to also include each cell's evaluated value alongside the formula in the JSON output. | Add support for a command‑line parameter that restricts the export to a single specified worksheet while keeping the same JSON structure.
// Common Searches: how to use Aspose.Cells in C# to export cell formulas to JSON | C# extract formulas with addresses from multiple sheets using Aspose.Cells | save Excel workbook formulas as JSON with sheet names in .NET | Aspose.Cells iterate cells and get formula property example | export only selected worksheet formulas to JSON using Aspose.Cells command line
// Tags: Aspose.Cells export formulas to JSON | C# extract cell formulas by sheet | serialize Excel formulas with sheet names .NET | iterate worksheet cells Aspose.Cells | command line worksheet filter Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates each worksheet and its cells, gathers non‑empty Formula properties together with cell addresses, groups them by worksheet name, and writes the resulting dictionary to an indented JSON file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "formulas.json";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Dictionary to hold sheet name -> list of formulas
            var sheetFormulas = new Dictionary<string, List<FormulaInfo>>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var formulas = new List<FormulaInfo>();
                Cells cells = sheet.Cells;

                // Iterate through all cells in the sheet
                foreach (Cell cell in cells)
                {
                    // A cell contains a formula if its Formula property is not empty
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        formulas.Add(new FormulaInfo
                        {
                            Address = cell.Name,   // e.g., "A1"
                            Formula = cell.Formula // the formula string
                        });
                    }
                }

                // Add to dictionary only if the sheet contains formulas
                if (formulas.Count > 0)
                {
                    sheetFormulas[sheet.Name] = formulas;
                }
            }

            // Serialize the dictionary to JSON with indentation
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string json = JsonSerializer.Serialize(sheetFormulas, jsonOptions);

            // Write JSON to file
            File.WriteAllText(outputPath, json);
            Console.WriteLine($"Formulas have been exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper class to represent a cell's formula information
    class FormulaInfo
    {
        public string? Address { get; set; }
        public string? Formula { get; set; }
    }
}

// Title: Aspose.Cells C# – Extract All Worksheet Formulas and Export to JSON
// Description: Loads an existing workbook, scans the first worksheet's used range, captures every cell that contains a formula, logs the address and formula, builds an address‑formula list, serializes it to pretty‑printed JSON, and writes the result to a file. Ideal for auditing, version control, or feeding formulas into custom processors.
// Keywords: Aspose.Cells extract formulas C# | export Excel formulas to JSON | list cell formulas Aspose.Cells | retrieve formula text .NET | save formulas as JSON file | Aspose.Cells example GitHub | C# spreadsheet automation | Excel formula extraction
// Common Searches: how to get formula text from each cell using Aspose.Cells | export Excel formulas to JSON with C# | iterate used range and collect formulas Aspose.Cells | Aspose.Cells example for extracting formulas | C# code to save worksheet formulas as JSON
// Developer Intent: Collect every formula in a worksheet and generate a JSON document that maps cell addresses to their formula strings.
// Use Cases: Create a read‑only documentation snapshot of all calculations for audit trails. | Maintain a version‑controlled list of formulas to detect changes between workbook revisions. | Supply extracted formulas to a custom engine that validates, transforms, or re‑calculates spreadsheet logic.
// AI Prompts: Generate C# code using Aspose.Cells that extracts all formulas from a worksheet and writes them to a formatted JSON file. | Provide a reusable method that returns a dictionary of cell addresses and formula strings for any Aspose.Cells worksheet. | Explain performance‑optimised techniques for extracting formulas from large Excel files and serializing them to JSON with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Cells;

// Loads an existing workbook, scans the first worksheet's used range, captures every cell that contains a formula, logs the address and formula, builds an address‑formula list, serializes it to pretty‑printed JSON, and writes the result to a file. Ideal for auditing, version control, or feeding formulas into custom processors.
class Program
{
    static void Main()
    {
        // Load the existing workbook (create rule)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Collect formula information
        var formulaList = new List<Dictionary<string, string>>();

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell != null && cell.IsFormula)
                {
                    string address = cell.Name;      // e.g., "A1"
                    string formula = cell.Formula;   // formula text

                    // Log to console
                    Console.WriteLine($"{address}: {formula}");

                    // Store for JSON export
                    formulaList.Add(new Dictionary<string, string>
                    {
                        ["Address"] = address,
                        ["Formula"] = formula
                    });
                }
            }
        }

        // Serialize the collected formulas to JSON
        string json = JsonSerializer.Serialize(formulaList, new JsonSerializerOptions { WriteIndented = true });

        // Export JSON to a file (save rule)
        string jsonPath = "formulas.json";
        File.WriteAllText(jsonPath, json);

        Console.WriteLine($"Formulas have been exported to {jsonPath}");
    }
}

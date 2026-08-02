// Title: Export Excel Formula Cells to CSV (Address, Formula, Result) – Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook with Aspose.Cells, forces formula calculation, scans the first worksheet for formula cells, and writes each cell's address, raw formula, and evaluated value to a UTF‑8 CSV file using RFC 4180‑compliant escaping.
// Keywords: Aspose.Cells | C# | .NET | export formulas to CSV | Excel formula extraction | cell address | formula evaluation | RFC 4180 | UTF-8 CSV | workbook.CalculateFormula
// Common Searches: Aspose.Cells export formula cells to CSV | C# list Excel formulas with results | how to get cell address and formula using Aspose | save Excel formulas as CSV .NET | extract calculated values from Excel with Aspose.Cells
// Developer Intent: Create a CSV file that lists every formula cell’s address, its original formula text, and the computed result from an Excel workbook using Aspose.Cells in C#.
// Use Cases: Audit all calculations in a financial model for compliance. | Provide downstream systems with a lightweight CSV of formulas and results. | Version‑control spreadsheet logic by exporting formulas to source‑friendly format. | Generate documentation of spreadsheet calculations for technical reviews.
// AI Prompts: Write C# code with Aspose.Cells that exports each formula cell to a CSV, handling commas, quotes, and newlines per RFC 4180. | Modify the sample to add a column for the worksheet name in the CSV output. | Add error handling for missing files and log the count of exported formula cells.

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// Loads an Excel workbook with Aspose.Cells, forces formula calculation, scans the first worksheet for formula cells, and writes each cell's address, raw formula, and evaluated value to a UTF‑8 CSV file using RFC 4180‑compliant escaping.
class ExportFormulaCsv
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Ensure all formulas are calculated
        workbook.CalculateFormula();

        // Prepare CSV content with header
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Address,Formula,Result");

        // Work with the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Iterate through all cells and collect formula information
        foreach (Cell cell in cells)
        {
            if (cell.IsFormula)
            {
                string address = cell.Name;               // e.g., "B2"
                string formula = cell.Formula;            // the formula text
                string result = cell.Value?.ToString() ?? string.Empty; // evaluated result

                // Escape fields that may contain commas or quotes
                address = EscapeCsv(address);
                formula = EscapeCsv(formula);
                result = EscapeCsv(result);

                csvBuilder.AppendLine($"{address},{formula},{result}");
            }
        }

        // Save the CSV file (replace with your desired output path)
        string outputPath = "formulas.csv";
        File.WriteAllText(outputPath, csvBuilder.ToString(), Encoding.UTF8);
    }

    // Helper to escape CSV fields according to RFC 4180
    static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}

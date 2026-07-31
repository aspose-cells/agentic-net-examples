// Title: Export Excel formulas with addresses and values to CSV using Aspose.Cells for .NET
// Description: Loads an Excel workbook, forces a full formula calculation, iterates through every worksheet, extracts each formula's A1 address, expression, and last evaluated value, escapes special characters, and writes the data to a CSV file with a header row.
// Keywords: Aspose.Cells | C# export formulas to CSV | Excel formula address | formula value extraction .NET | CSV escaping Aspose | calculate formulas before export | extract Excel formulas programmatically
// Common Searches: how to export all formulas from Excel to CSV using Aspose.Cells | C# code to write cell address, formula, and result to CSV | Aspose.Cells iterate cells and get formula text and value | export Excel formula results with proper CSV escaping | Aspose.Cells export formulas with worksheet name
// Developer Intent: Generate a CSV file that lists every formula in an Excel workbook together with its cell address and the most recent calculated result.
// Use Cases: Audit financial models by creating a snapshot of all formulas, locations, and current results. | Supply downstream systems that accept only CSV with a feed of calculated values from complex workbooks. | Track changes between workbook versions by comparing exported formula/value snapshots.
// AI Prompts: Write C# code using Aspose.Cells to export each formula's address, expression, and evaluated value to a CSV file, handling commas, quotes, and newlines correctly. | Extend the sample to include the worksheet name as an additional column in the CSV output. | Add comprehensive error handling for missing files, permission issues, and cells that fail to export while logging problems and continuing the process.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, forces a full formula calculation, iterates through every worksheet, extracts each formula's A1 address, expression, and last evaluated value, escapes special characters, and writes the data to a CSV file with a header row.
class ExportFormulasToCsv
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Path where the CSV will be saved
        string csvPath = "formulas.csv";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(excelPath);

        // Calculate all formulas to ensure values are up‑to‑date
        workbook.CalculateFormula();

        // Open a StreamWriter for the CSV output
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Write CSV header
            writer.WriteLine("Address,Formula,Value");

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells that contain data
                foreach (Cell cell in cells)
                {
                    // Process only cells that have a formula
                    if (cell.IsFormula)
                    {
                        // Cell address in A1 notation (e.g., "B2")
                        string address = cell.Name;

                        // Formula text as stored in the cell
                        string formula = cell.Formula;

                        // Last evaluated value of the formula
                        string value = cell.Value?.ToString() ?? string.Empty;

                        // Escape fields that may contain commas or quotes
                        formula = EscapeCsv(formula);
                        value = EscapeCsv(value);

                        // Write a CSV line
                        writer.WriteLine($"{address},{formula},{value}");
                    }
                }
            }
        }

        Console.WriteLine($"Formulas exported successfully to '{csvPath}'.");
    }

    // Helper method to escape CSV fields containing commas, quotes, or newlines
    static string EscapeCsv(string field)
    {
        if (field.Contains("\""))
            field = field.Replace("\"", "\"\"");

        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
            field = $"\"{field}\"";

        return field;
    }
}

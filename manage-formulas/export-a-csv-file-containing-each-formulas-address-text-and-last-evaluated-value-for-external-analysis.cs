// Title: Export Excel formulas with addresses and evaluated values to CSV using Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook, forces full formula calculation, then iterates all worksheets to capture each formula cell’s A1 address, raw formula text, and the last calculated value. The information is CSV‑escaped and saved to "formulas.csv".
// Keywords: Aspose.Cells | C# | .NET | export formulas to CSV | Excel formula address | last calculated value | extract cell formulas | calculate workbook formulas | Excel automation | CSV dump of formulas
// Common Searches: Aspose.Cells export all formulas to CSV C# | How to get formula address and value from Excel using Aspose.Cells | Write Excel formulas and results to a CSV file in .NET | Calculate workbook formulas before exporting with Aspose.Cells | Extract formula metadata from multiple worksheets C#
// Developer Intent: Generate a CSV file that lists every formula in an Excel workbook together with its cell address and the most recent evaluated result.
// Use Cases: Audit financial models by providing a complete formula inventory with results. | Supply external reporting or analytics systems with formula metadata and outcomes. | Create a snapshot of spreadsheet calculations for performance benchmarking. | Facilitate regulatory compliance checks that require documented formula logic. | Assist in migrating complex spreadsheets to other platforms by exporting formula details.
// AI Prompts: Add a column for the worksheet name to each CSV row while preserving address, formula, and value. | Implement robust error handling for missing input files and log cells where value conversion fails. | Write a routine that reads the generated CSV and groups formulas by worksheet for further analysis. | Create a version of the exporter that writes JSON instead of CSV for API consumption. | Enhance the code to skip hidden rows/columns and only export visible formula cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx workbook, forces full formula calculation, then iterates all worksheets to capture each formula cell’s A1 address, raw formula text, and the last calculated value. The information is CSV‑escaped and saved to "formulas.csv".
class ExportFormulasToCsv
{
    static void Main()
    {
        // Load the workbook (create/load lifecycle)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Calculate all formulas so that each cell has its last evaluated value
        workbook.CalculateFormula();

        // Prepare CSV file for output
        string csvPath = "formulas.csv";
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
                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // Cell address in A1 notation
                        string address = cell.Name;

                        // Formula text as stored in the cell
                        string formula = cell.Formula;

                        // Last evaluated value (converted to string, handling nulls)
                        string value = cell.Value?.ToString() ?? string.Empty;

                        // Escape commas and quotes for CSV compliance
                        string escapedFormula = $"\"{formula.Replace("\"", "\"\"")}\"";
                        string escapedValue = $"\"{value.Replace("\"", "\"\"")}\"";

                        // Write a CSV line
                        writer.WriteLine($"{address},{escapedFormula},{escapedValue}");
                    }
                }
            }
        }

        Console.WriteLine($"Formulas have been exported to '{csvPath}'.");
    }
}

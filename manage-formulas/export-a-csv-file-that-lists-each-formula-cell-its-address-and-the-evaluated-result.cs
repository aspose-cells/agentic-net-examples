// Title: Export Excel formula cells with addresses and evaluated results to CSV using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, forces a full recalculation, scans the used range of each worksheet, captures every cell that contains a formula, records its A1‑style address and displayed value, applies CSV escaping, and writes the data to a UTF‑8 CSV file with a header.
// Keywords: Aspose.Cells | .NET | C# | export formulas to CSV | list formula cells | cell address | evaluated result | calculate Excel formulas | Excel to CSV conversion | extract formula values
// Common Searches: Aspose.Cells export formula cells to CSV C# | how to list Excel formulas with results using Aspose | C# code to write formula addresses and values to CSV | extract calculated values from workbook with Aspose.Cells | save Excel formula results as CSV file
// Developer Intent: Generate a CSV report that enumerates each formula cell’s address and its current calculated value.
// Use Cases: Create an audit trail of all formulas and their outcomes for compliance reviews. | Provide downstream systems with a lightweight CSV of calculated results without sharing the full workbook. | Offer business users a quick reference of which cells contain formulas and what they evaluate to.
// AI Prompts: Write C# code that iterates through all worksheets in a workbook and exports every formula cell’s address and result to a CSV using Aspose.Cells. | Explain how to correctly escape commas and double quotes in CSV output when the formula result contains special characters. | Show how to extend the example to also include the original formula text (e.g., "=A1+B1") alongside the address and result.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Loads an Excel workbook, forces a full recalculation, scans the used range of each worksheet, captures every cell that contains a formula, records its A1‑style address and displayed value, applies CSV escaping, and writes the data to a UTF‑8 CSV file with a header.
class ExportFormulaCsv
{
    static void Main()
    {
        // Input Excel file path
        string inputPath = "input.xlsx";

        // Output CSV file path
        string outputPath = "formulas.csv";

        // Load the workbook (create rule)
        Workbook workbook = new Workbook(inputPath);

        // Calculate all formulas in the workbook (ensure results are up‑to‑date)
        workbook.CalculateFormula();

        // Access the first worksheet (you can modify to iterate all worksheets if needed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Prepare CSV content with a header line
        StringBuilder csvBuilder = new StringBuilder();
        csvBuilder.AppendLine("Address,Result");

        // Determine the used range to limit iteration
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Iterate through each cell in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell != null && cell.IsFormula)
                {
                    // Cell address (e.g., A1)
                    string address = cell.Name;

                    // Evaluated result as a string (formatted as it appears in Excel)
                    string result = cell.StringValue;

                    // Simple CSV escaping: double quotes are escaped by doubling them,
                    // and the field is wrapped in quotes if it contains a comma or quote.
                    if (result.Contains("\""))
                        result = result.Replace("\"", "\"\"");

                    if (result.Contains(",") || result.Contains("\""))
                        result = $"\"{result}\"";

                    csvBuilder.AppendLine($"{address},{result}");
                }
            }
        }

        // Write the CSV content to the output file (save rule)
        File.WriteAllText(outputPath, csvBuilder.ToString(), Encoding.UTF8);

        Console.WriteLine($"Export completed. CSV saved to '{outputPath}'.");
    }
}

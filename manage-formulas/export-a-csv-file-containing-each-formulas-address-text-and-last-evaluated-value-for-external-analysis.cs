// Title: Export each formula’s sheet‑qualified address, expression, and last calculated value to a CSV file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel workbook with Aspose.Cells, scans all worksheets, and writes every cell containing a formula to a CSV file, including the sheet name, cell address, formula string, and the cell’s last evaluated value. | Add a helper method that correctly escapes commas, double quotes, and line breaks for each CSV field when exporting formula data from Aspose.Cells. | Show how to build a full cell reference that combines the worksheet name and the cell name (e.g., Sheet1!A1) while iterating cells with Aspose.Cells. | Demonstrate writing a CSV header and appending rows for formula address, formula text, and value using StreamWriter in a .NET console application.
// Common Searches: Aspose.Cells C# export formulas with values to CSV file | How to get formula text and last calculated result from each cell using Aspose.Cells | C# write Excel formula address and result to CSV with proper escaping | Iterate all cells in a workbook and output formulas with sheet name Aspose.Cells | CSV field escaping for commas and quotes when exporting Excel data in .NET
// Tags: Aspose.Cells formula export CSV | retrieve formula text and evaluated value .NET | CSV escaping commas quotes newlines C# | sheet‑qualified cell address Aspose.Cells | iterate worksheets and cells Aspose.Cells example

using Aspose.Cells;
using System;
using System.IO;

// The example loads an input.xlsx workbook, walks through every worksheet and cell, and writes each formula‑containing cell to formulas.csv. Each CSV row records the sheet‑qualified address (e.g., Sheet1!A1), the formula string, and the last evaluated value, with proper CSV escaping for commas, quotes, and newlines.
class Program
{
    static void Main()
    {
        // Load the workbook (adjust the file path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Create a CSV file to store formula information
        using (StreamWriter csvWriter = new StreamWriter("formulas.csv"))
        {
            // Write CSV header
            csvWriter.WriteLine("Address,Formula,Value");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Iterate through all cells in the worksheet
                foreach (Cell cell in cells)
                {
                    // Check if the cell contains a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Build the full address (including sheet name)
                        string address = $"{sheet.Name}!{cell.Name}";

                        // Retrieve the formula text
                        string formula = cell.Formula;

                        // Retrieve the last evaluated value (as string)
                        string value = cell.Value?.ToString() ?? string.Empty;

                        // Escape fields for CSV compliance
                        address = EscapeCsv(address);
                        formula = EscapeCsv(formula);
                        value = EscapeCsv(value);

                        // Write the record to the CSV file
                        csvWriter.WriteLine($"{address},{formula},{value}");
                    }
                }
            }
        }
    }

    // Helper method to escape commas, quotes, and newlines in CSV fields
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

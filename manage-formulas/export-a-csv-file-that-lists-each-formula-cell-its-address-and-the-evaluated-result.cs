// Title: Export formula cells with addresses and evaluated results to a CSV file using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook with Aspose.Cells, forces calculation of all formulas, and creates a CSV listing each formula cell’s address, the formula text, and its computed value. | Add logic to the Aspose.Cells example to correctly escape commas, double quotes, and line breaks when writing formula data to a CSV file. | Modify the sample so the output CSV path is supplied via a command‑line argument and include a header row in the exported file.
// Common Searches: Aspose.Cells .NET generate CSV of formula cells with address and result | C# list all formulas and their calculated values from an Excel workbook | How to write Excel formula cells to CSV using Aspose.Cells and include cell references | Escaping commas and quotes when exporting Aspose.Cells data to CSV in C# | Pass output file name to Aspose.Cells script that exports formulas
// Tags: generate CSV of formulas Aspose.Cells | calculate all workbook formulas Aspose.Cells | retrieve cell address and formula text Aspose.Cells | CSV escaping for commas and quotes C# | iterate used rows and columns Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, forces formula calculation, iterates over the used range, and writes each formula cell's address, formula string, and evaluated result to a CSV file with proper escaping.
class ExportFormulaCsv
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are calculated
        workbook.CalculateFormula();

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Prepare the CSV output file
        using (StreamWriter writer = new StreamWriter("formulas.csv"))
        {
            // Write CSV header
            writer.WriteLine("Address,Formula,Result");

            // Iterate through all used rows and columns
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        string address = cell.Name;          // e.g., "B2"
                        string formula = cell.Formula;       // the formula string
                        string result = cell.Value?.ToString() ?? string.Empty; // evaluated result

                        // Escape commas and double quotes for CSV compliance
                        formula = EscapeForCsv(formula);
                        result = EscapeForCsv(result);

                        // Write the line to CSV
                        writer.WriteLine($"{address},{formula},{result}");
                    }
                }
            }
        }

        Console.WriteLine("CSV export completed.");
    }

    // Helper method to escape fields containing commas or quotes
    private static string EscapeForCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}

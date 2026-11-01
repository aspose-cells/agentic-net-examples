// Title: Export formula cells with their addresses and evaluated results to a CSV file using Aspose.Cells for .NET
// Common Searches: how to generate a CSV of all formula cells with addresses and results using Aspose.Cells in C# | C# Aspose.Cells export worksheet name, cell address, formula and calculated value to CSV | list Excel formula cells and evaluated results to a CSV file with the Aspose.Cells .NET library
// Tags: Aspose.Cells export formula cells CSV | Aspose.Cells retrieve cell address and value .NET | Aspose.Cells iterate worksheets to extract formulas C# | Aspose.Cells calculate workbook formulas before export | CSV field escaping for Excel data Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, forces calculation of all formulas, walks through each worksheet and cell, captures the worksheet name, cell address, formula text, and the evaluated result for every formula cell, applies RFC‑4180 CSV escaping, and writes the information to a CSV file with a header row.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the CSV will be saved
        string outputCsvPath = "formula_cells.csv";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;

        try
        {
            // Load the workbook (no password supplied)
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // Detect password‑protected file via exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine($"Error: The workbook \"{inputPath}\" is password protected.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
            return;
        }

        try
        {
            // Ensure all formulas are evaluated
            workbook.CalculateFormula();

            // Open a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(outputCsvPath))
            {
                // Write CSV header
                writer.WriteLine("Worksheet,Address,Formula,Value");

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all cells in the worksheet
                    foreach (Cell cell in cells)
                    {
                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            // Get cell address (e.g., A1)
                            string address = cell.Name;

                            // Get the formula string
                            string formula = cell.Formula;

                            // Get the evaluated result as a string
                            string value = cell.Value?.ToString() ?? string.Empty;

                            // Escape fields that may contain commas or quotes
                            address = EscapeCsv(address);
                            formula = EscapeCsv(formula);
                            value = EscapeCsv(value);

                            // Write a line to the CSV file
                            writer.WriteLine($"{sheet.Name},{address},{formula},{value}");
                        }
                    }
                }
            }

            Console.WriteLine($"CSV file \"{outputCsvPath}\" has been created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }

    // Helper method to escape CSV fields according to RFC 4180
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

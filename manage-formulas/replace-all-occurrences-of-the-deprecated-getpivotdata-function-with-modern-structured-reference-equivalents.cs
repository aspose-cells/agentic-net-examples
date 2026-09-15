// Title: Replace deprecated GETPIVOTDATA formulas with structured references in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, scans every cell for a GETPIVOTDATA formula, and rewrites it as a structured reference. | Create a helper method that parses a GETPIVOTDATA string and returns the equivalent structured‑reference syntax for use in Aspose.Cells formulas. | Show how to save the workbook after converting all GETPIVOTDATA formulas to structured references with Aspose.Cells.
// Common Searches: how to convert GETPIVOTDATA to structured reference using Aspose.Cells C# | Aspose.Cells programmatically replace deprecated GETPIVOTDATA formulas | C# iterate through Excel cells and update GETPIVOTDATA to structured reference
// Tags: replace GETPIVOTDATA with structured reference Aspose.Cells | scan and modify Excel formulas C# | convert deprecated Excel function Aspose.Cells .NET | bulk formula update .xlsx C# | structured reference generation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an input .xlsx workbook with Aspose.Cells, iterates through all worksheets and cells, detects formulas containing the GETPIVOTDATA function, converts each to a simple structured‑reference syntax via a placeholder parser, replaces the original formula, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Scan all cells within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Identify formulas that contain the deprecated GETPIVOTDATA function
                        if (cell.IsFormula && cell.Formula.IndexOf("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string oldFormula = cell.Formula;

                            // Convert the GETPIVOTDATA call to a modern structured reference
                            string newFormula = ConvertGetPivotDataToStructuredReference(oldFormula);

                            // Replace the formula with the new structured reference
                            cell.Formula = newFormula;
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple placeholder conversion: extracts the field name from GETPIVOTDATA and builds a structured reference.
    // Real-world scenarios would require full parsing of all arguments.
    static string ConvertGetPivotDataToStructuredReference(string formula)
    {
        // Locate the GETPIVOTDATA function name
        int funcPos = formula.IndexOf("GETPIVOTDATA", StringComparison.OrdinalIgnoreCase);
        if (funcPos < 0) return formula;

        // Find the first quoted field name argument
        int firstQuote = formula.IndexOf('\"', funcPos);
        int secondQuote = formula.IndexOf('\"', firstQuote + 1);
        if (firstQuote < 0 || secondQuote < 0) return formula;

        // Extract the field name (e.g., "Sum of Sales")
        string fieldName = formula.Substring(firstQuote + 1, secondQuote - firstQuote - 1);

        // Build a structured reference using the extracted field name
        // Example: =GETPIVOTDATA("Sum of Sales",$A$3) -> =[@[Sum of Sales]]
        string structuredReference = $"[@[{fieldName}]]";

        // Return the new formula (as a simple reference; adjust as needed for complex cases)
        return structuredReference;
    }
}

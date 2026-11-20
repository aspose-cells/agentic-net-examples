// Title: Create a volatile‑function audit report after recalculating formulas with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# list cells containing volatile functions after CalculateFormula | how to generate a volatile function report in an Excel workbook using Aspose.Cells .NET | C# identify and export volatile Excel formulas to a new worksheet with Aspose.Cells | retrieve addresses of cells with NOW, RAND, OFFSET functions after workbook recalculation in Aspose.Cells
// Tags: volatile function detection Aspose.Cells C# | list volatile cells after CalculateFormula | create audit worksheet for volatile formulas .NET | extract volatile formula addresses Aspose.Cells | Excel volatile functions audit using Aspose

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel file, runs Workbook.CalculateFormula to recalculate all formulas, scans each worksheet for cells whose formulas start with known volatile functions (e.g., NOW, RAND, OFFSET), collects their addresses in "SheetName!A1" format, prints them, adds a new worksheet named "VolatileReport" with the list, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_with_volatile_report.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook; if the file is password‑protected, an exception will be thrown
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Collection to store addresses of cells that contain volatile functions
            List<string> volatileCells = new List<string>();

            // Set of known volatile function names (case‑insensitive)
            HashSet<string> volatileFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL", "AREAS", "GETPIVOTDATA"
            };

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range to limit iteration
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only cells that contain a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula;

                            // Remove leading '=' if present
                            if (formula.StartsWith("="))
                                formula = formula.Substring(1);

                            // Extract the function name (text before the first '(' )
                            int parenIndex = formula.IndexOf('(');
                            if (parenIndex > 0)
                            {
                                string functionName = formula.Substring(0, parenIndex).Trim();

                                // Check if the function is in the volatile list
                                if (volatileFunctions.Contains(functionName))
                                {
                                    // Store address in the format SheetName!A1
                                    volatileCells.Add($"{sheet.Name}!{cell.Name}");
                                }
                            }
                        }
                    }
                }
            }

            // Output the list to the console
            Console.WriteLine("Volatile cells after workbook recalculation:");
            foreach (string address in volatileCells)
            {
                Console.WriteLine(address);
            }

            // Optionally create a report sheet with the results
            int reportIndex = workbook.Worksheets.Add();
            Worksheet reportSheet = workbook.Worksheets[reportIndex];
            reportSheet.Name = "VolatileReport";
            reportSheet.Cells[0, 0].PutValue("Volatile Cells");
            for (int i = 0; i < volatileCells.Count; i++)
            {
                reportSheet.Cells[i + 1, 0].PutValue(volatileCells[i]);
            }

            // Save the workbook (including the new report sheet)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}

// Title: Validate Excel array formula syntax with Aspose.Cells for .NET before calling Workbook.CalculateFormula
// Common Searches: how to detect invalid array formulas in an Excel file using Aspose.Cells C# | skip Workbook.CalculateFormula when array formula errors exist in Aspose.Cells | programmatically verify array formula syntax in .NET before calculating workbook | Aspose.Cells C# iterate cells to check array formula validity | prevent runtime errors from bad array formulas with Aspose.Cells calculation
// Tags: array formula error detection Aspose.Cells | cell.Calculate exception handling C# | Workbook.CalculateFormula prevalidation Aspose.Cells | scan worksheets for array formulas C# | Aspose.Cells CalculationOptions usage

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, scans each worksheet's used range, identifies cells marked as array formulas, attempts to calculate each cell with CalculationOptions, logs any exceptions as invalid formulas, aborts the overall Workbook.CalculateFormula when invalid array formulas are found, and finally saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            bool hasInvalidArrayFormula = false;
            var calcOptions = new CalculationOptions(); // options for cell calculation

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range limits
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Scan each cell within the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check if the cell contains an array formula
                        if (cell.IsArrayFormula)
                        {
                            // Retrieve the raw array formula string (including braces)
                            string arrayFormula = cell.Formula;

                            // Attempt to evaluate the formula; treat any exception as invalid
                            try
                            {
                                cell.Calculate(calcOptions);
                            }
                            catch (Exception)
                            {
                                hasInvalidArrayFormula = true;
                                Console.WriteLine(
                                    $"Invalid array formula detected at {sheet.Name}!{cell.Name}: {arrayFormula}");
                            }
                        }
                    }
                }
            }

            // Perform calculation only if no invalid array formulas were found
            if (!hasInvalidArrayFormula)
            {
                try
                {
                    workbook.CalculateFormula();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during workbook calculation: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Workbook calculation aborted due to invalid array formulas.");
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}

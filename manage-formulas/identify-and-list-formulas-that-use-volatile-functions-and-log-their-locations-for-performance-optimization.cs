// Title: Identify and export volatile Excel formulas with their sheet‑cell addresses using Aspose.Cells for .NET
// Common Searches: C# code to list cells that contain volatile functions like NOW or RAND using Aspose.Cells | how to generate a report of volatile formulas in an .xlsx workbook with Aspose.Cells .NET | scan an Excel workbook for volatile formulas and save sheet!cell locations to a text file in C# | Aspose.Cells find and log volatile functions for performance optimization in .NET applications | enumerate cells with volatile Excel formulas and export their addresses using the Aspose.Cells API
// Tags: volatile formula detection with Aspose.Cells | export volatile formula locations to text file | scan workbook for volatile functions C# | Aspose.Cells performance optimization of formulas | list sheet cell addresses of volatile Excel functions

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace VolatileFormulaScanner
{
    // The example loads an Excel workbook with Aspose.Cells, iterates each worksheet's used range, checks every formula cell against a predefined list of volatile functions (e.g., NOW, RAND, OFFSET), records matching cells as "SheetName!CellAddress", prints the locations to the console, and writes them to a text report for performance analysis.
    class Program
    {
        // List of volatile functions in Excel (case‑insensitive)
        private static readonly string[] VolatileFunctions = new[]
        {
            "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "GETPIVOTDATA()", "RTD()", "AREAS()", "CHOOSE()", "HYPERLINK()",
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL", "GETPIVOTDATA", "RTD", "AREAS", "CHOOSE", "HYPERLINK"
        };

        static void Main(string[] args)
        {
            try
            {
                // Path to the Excel file to be analyzed
                string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (using Aspose.Cells)
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                // Collection to store locations of volatile formulas
                List<string> volatileFormulaLocations = new List<string>();

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet to limit iteration
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                    // If the worksheet is empty, skip it
                    if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Iterate through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Check if the cell contains a formula
                            if (cell.IsFormula)
                            {
                                string formula = cell.Formula;

                                // Simple detection: check if any volatile function name appears in the formula string (case‑insensitive)
                                foreach (string volatileFunc in VolatileFunctions)
                                {
                                    if (formula.IndexOf(volatileFunc, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        // Record the location: SheetName!CellName
                                        string location = $"{sheet.Name}!{cell.Name}";
                                        volatileFormulaLocations.Add(location);
                                        // No need to check other volatile functions for this cell
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Volatile formulas found at the following locations:");
                foreach (string loc in volatileFormulaLocations)
                {
                    Console.WriteLine(loc);
                }

                // Optionally, write the locations to a text file for further analysis
                string outputPath = "VolatileFormulasReport.txt";
                try
                {
                    File.WriteAllLines(outputPath, volatileFormulaLocations);
                    Console.WriteLine($"Report saved to: {outputPath}");
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write report: {writeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

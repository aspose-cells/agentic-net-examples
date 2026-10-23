// Title: Generate a report of cells containing volatile Excel functions after recalculating formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells detect volatile functions after workbook.CalculateFormula | list Excel cells that use NOW() TODAY() RAND() with Aspose.Cells .NET | create a text file report of volatile formulas in an Excel workbook using Aspose.Cells | how to enumerate cells containing OFFSET or INDIRECT after recalculating formulas in C# | programmatically identify volatile Excel functions in each worksheet with Aspose.Cells
// Tags: Aspose.Cells volatile function detection | C# enumerate Excel cells with volatile formulas | generate text report of volatile functions | recalculate workbook and scan formulas Aspose.Cells | Excel volatile functions NOW TODAY RAND detection

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace VolatileFunctionsReport
{
    // The example loads an Excel file with Aspose.Cells, forces a full formula recalculation, scans every used cell for known volatile functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL), records each matching cell address with its worksheet name, and writes a timestamped text report listing all identified volatile cells or indicating none were found.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the input workbook
                string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // List to hold addresses of cells containing volatile functions
                List<string> volatileCells = new List<string>();

                // Known volatile function names (case‑insensitive)
                string[] volatileFunctions = new[]
                {
                    "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()"
                };

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null) continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = startRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = startCol + usedRange.ColumnCount - 1;

                    // Iterate through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Check if the cell contains a formula
                            string formula = cell.Formula;
                            if (string.IsNullOrEmpty(formula)) continue;

                            // Determine if the formula uses any volatile function
                            if (volatileFunctions.Any(v => formula.IndexOf(v, StringComparison.OrdinalIgnoreCase) >= 0))
                            {
                                // Record the cell address with sheet name
                                string address = $"{sheet.Name}!{cell.Name}";
                                volatileCells.Add(address);
                            }
                        }
                    }
                }

                // Prepare the report content
                string reportPath = "volatile_cells_report.txt";
                try
                {
                    using (StreamWriter writer = new StreamWriter(reportPath))
                    {
                        writer.WriteLine("Volatile Functions Report");
                        writer.WriteLine($"Generated on: {DateTime.Now}");
                        writer.WriteLine();

                        if (volatileCells.Count == 0)
                        {
                            writer.WriteLine("No volatile functions were found in the workbook.");
                        }
                        else
                        {
                            writer.WriteLine("Cells containing volatile functions:");
                            foreach (string addr in volatileCells)
                            {
                                writer.WriteLine(addr);
                            }
                        }
                    }

                    Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to write report: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Identify and list Excel cells that contain volatile functions using Aspose.Cells for .NET
// Common Searches: C# code to find cells with volatile formulas like NOW or RAND using Aspose.Cells | how to enumerate addresses of volatile function cells in an .xlsx with Aspose.Cells .NET | detect and list volatile Excel functions in a workbook via Aspose.Cells API | retrieve A1 style addresses of cells that use OFFSET or INDIRECT in C# Aspose.Cells | scan all worksheets for volatile formulas with Aspose.Cells for .NET
// Tags: volatile function detection Aspose.Cells .NET | list cells with volatile formulas C# | enumerate A1 addresses of formula cells Aspose.Cells | scan workbook used range for volatile functions | identify volatile Excel functions in .xlsx using Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace VolatileFormulaFinder
{
    // The example loads an .xlsx workbook with Aspose.Cells for .NET, defines the set of Excel volatile functions, iterates each worksheet’s used range, checks every formula cell for those functions (case‑insensitive), records the sheet name and A1‑style address of matching cells, and outputs the list of volatile‑function cell addresses.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be analyzed
            string workbookPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found - {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                Workbook workbook = new Workbook(workbookPath);

                // List of volatile Excel functions (case-insensitive)
                string[] volatileFunctions = new string[]
                {
                    "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "AREAS()", "GETPIVOTDATA()", "RTD()",
                    "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL", "AREAS", "GETPIVOTDATA", "RTD"
                };

                // Store addresses of cells that contain volatile functions
                List<string> volatileCellAddresses = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet to limit iteration
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                    if (usedRange == null) continue;

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

                                // Determine if the formula uses any volatile function
                                foreach (string volatileFunc in volatileFunctions)
                                {
                                    // Simple containment check (case-insensitive)
                                    if (formula.IndexOf(volatileFunc, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        // Build address in A1 style with sheet name
                                        string address = $"{sheet.Name}!{cell.Name}";
                                        volatileCellAddresses.Add(address);
                                        break; // No need to check other volatile functions for this cell
                                    }
                                }
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Cells containing volatile functions:");
                foreach (string addr in volatileCellAddresses)
                {
                    Console.WriteLine(addr);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // Optional: keep console window open when running outside debugger
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}

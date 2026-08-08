// Title: C# – Find and list Excel cells that use volatile functions with Aspose.Cells
// Description: Loads a workbook, optionally enables the calculation chain, evaluates all formulas, then iterates the used range of every worksheet. Each formula cell is inspected for known volatile functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL). Matching cells are recorded with their sheet‑qualified A1 address, printed to the console, and the workbook is saved.
// Keywords: Aspose.Cells volatile function detection | C# find volatile formulas in Excel | list cells with NOW TODAY RAND Aspose | Excel performance audit .NET | enumerate volatile formulas Aspose.Cells | search Excel formulas for volatile functions | Aspose.Cells calculation chain example
// Common Searches: how to locate volatile formulas in an Excel file using Aspose.Cells | C# code to list cells containing OFFSET or INDIRECT functions | find all NOW/TODAY formulas in a workbook with Aspose.Cells | detect volatile functions in Excel for .NET developers | Aspose.Cells example to audit volatile formulas
// Developer Intent: Identify every cell that contains a volatile Excel function and output its address.
// Use Cases: Performance review: flag volatile formulas that may slow recalculation. | Compliance reporting: generate an inventory of volatile functions for governance. | Automated refactoring: replace or log volatile formulas before publishing a workbook.
// AI Prompts: Create a reusable method that returns a List<string> of sheet‑qualified addresses for cells with any volatile function. | Adapt the sample to skip cells that belong to chart data ranges while still detecting volatile formulas. | Extend the code to write each volatile formula and its address to a CSV log file.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsVolatileFormulaFinder
{
    // Loads a workbook, optionally enables the calculation chain, evaluates all formulas, then iterates the used range of every worksheet. Each formula cell is inspected for known volatile functions (NOW, TODAY, RAND, RANDBETWEEN, OFFSET, INDIRECT, INFO, CELL). Matching cells are recorded with their sheet‑qualified A1 address, printed to the console, and the workbook is saved.
    class Program
    {
        static void Main()
        {
            // ---------- Create or Load Workbook ----------
            // Replace "input.xlsx" with the path to your workbook.
            Workbook workbook = new Workbook("input.xlsx");   // load existing workbook

            // Enable calculation chain to ensure formulas are evaluated (optional but useful)
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;
            workbook.CalculateFormula(); // calculate all formulas

            // List of known volatile functions (add more if needed)
            string[] volatileFunctions = new string[]
            {
                "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL"
            };

            // Collection to store addresses of cells containing volatile formulas
            List<string> volatileCellAddresses = new List<string>();

            // Iterate through all worksheets and cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                // Use the MaxDataRow/MaxDataColumn to limit iteration to used range
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // Check if the cell contains a formula
                        if (cell.IsFormula)
                        {
                            string formula = cell.Formula?.ToUpperInvariant() ?? string.Empty;
                            // Determine if any volatile function appears in the formula
                            foreach (string vf in volatileFunctions)
                            {
                                if (formula.Contains(vf))
                                {
                                    // Record the address in A1 style with sheet name
                                    string address = $"{sheet.Name}!{cell.Name}";
                                    volatileCellAddresses.Add(address);
                                    break; // No need to check other functions for this cell
                                }
                            }
                        }
                    }
                }
            }

            // Output the results
            Console.WriteLine("Cells with volatile functions:");
            foreach (string addr in volatileCellAddresses)
            {
                Console.WriteLine(addr);
            }

            // ---------- Save Workbook ----------
            // Replace "output.xlsx" with the desired output path.
            workbook.Save("output.xlsx");
        }
    }
}

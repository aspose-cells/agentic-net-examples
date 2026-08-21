// Title: Detect and Correct Formulas that Reference Deleted Rows using Aspose.Cells for .NET (C#)
// Description: A complete C# example that creates a workbook, adds formulas referencing specific rows, identifies all dependent cells with Cells.GetDependents, deletes the target rows with DeleteOptions.UpdateReference, and reports any formulas that turn into #REF! errors for manual correction.
// Keywords: Aspose.Cells GetDependents C# | update formula references after row deletion | detect #REF! errors Aspose.Cells | delete rows with reference update .NET | Excel formula dependency check | C# Aspose.Cells row removal | automatic formula adjustment
// Common Searches: How to find cells that depend on a row using Aspose.Cells C# | C# delete rows and keep formulas correct with Aspose.Cells | Identify #REF! after removing rows in Excel via Aspose.Cells | Get dependent cells before deleting rows Aspose.Cells .NET | Update formula references automatically when rows are removed
// Developer Intent: Locate formulas that point to rows slated for deletion, remove those rows while letting Aspose.Cells adjust references, and flag any resulting #REF! formulas for review.
// Use Cases: Iterate over columns of each row to be removed and call Cells.GetDependents to collect dependent cells. | Delete rows with Cells.DeleteRows using DeleteOptions.UpdateReference = true so formulas are auto‑adjusted. | After deletion, enumerate the previously collected cells, output their new formulas, and suggest manual fixes for any formula containing #REF!.
// AI Prompts: Generate a C# method that receives a Worksheet and a list of row indexes, returns all cells whose formulas reference those rows using Aspose.Cells. | Write code that deletes specified rows with DeleteOptions.UpdateReference enabled and logs formulas that become #REF! after the operation. | Explain how Aspose.Cells rewrites formula references when rows are removed and show how to verify the changes programmatically.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaReferenceChecker
{
    // A complete C# example that creates a workbook, adds formulas referencing specific rows, identifies all dependent cells with Cells.GetDependents, deletes the target rows with DeleteOptions.UpdateReference, and reports any formulas that turn into #REF! errors for manual correction.
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook and fill it with sample data
            // ------------------------------------------------------------
            Workbook wb = new Workbook();                     // create workbook
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate rows 1‑10 with simple numeric values
            for (int r = 0; r < 10; r++)
            {
                for (int c = 0; c < 3; c++)
                {
                    cells[r, c].PutValue(r + 1 + c * 0.1); // e.g., 1.0, 1.1, 1.2 …
                }
            }

            // Add formulas that reference rows 3 and 5 (zero‑based indexes 2 and 4)
            cells["D3"].Formula = "=A3+B3";   // references row 3
            cells["E5"].Formula = "=SUM(A5:C5)"; // references row 5
            cells["F7"].Formula = "=D3*E5";   // indirect reference to rows 3 & 5

            // ------------------------------------------------------------
            // 2. Determine which formulas reference the rows that will be deleted
            // ------------------------------------------------------------
            // Rows to delete (zero‑based): 2 (row 3) and 4 (row 5)
            int[] rowsToDelete = { 2, 4 };
            var dependentCells = new HashSet<Cell>(); // avoid duplicates

            foreach (int delRow in rowsToDelete)
            {
                // Scan all columns in the row to find cells that might be referenced
                int maxCol = cells.MaxDataColumn;
                for (int col = 0; col <= maxCol; col++)
                {
                    // Get cells that depend on the current cell (delRow, col)
                    // isAll = false – we only need references inside the same worksheet
                    Cell[] deps = cells.GetDependents(false, delRow, col);
                    if (deps != null)
                    {
                        foreach (Cell dep in deps)
                        {
                            dependentCells.Add(dep);
                        }
                    }
                }
            }

            // ------------------------------------------------------------
            // 3. Output the formulas that currently reference the soon‑to‑be‑deleted rows
            // ------------------------------------------------------------
            Console.WriteLine("Formulas that reference rows to be deleted (before deletion):");
            foreach (Cell dep in dependentCells)
            {
                Console.WriteLine($"{dep.Name}: {dep.Formula}");
            }

            // ------------------------------------------------------------
            // 4. Delete the rows with reference updating enabled
            // ------------------------------------------------------------
            DeleteOptions delOptions = new DeleteOptions
            {
                UpdateReference = true   // let Aspose.Cells adjust formulas automatically
            };

            // Delete rows in descending order to keep indexes valid
            Array.Sort(rowsToDelete);
            Array.Reverse(rowsToDelete);
            foreach (int delRow in rowsToDelete)
            {
                cells.DeleteRows(delRow, 1, delOptions);
            }

            // ------------------------------------------------------------
            // 5. After deletion, show the updated formulas and flag any #REF! errors
            // ------------------------------------------------------------
            Console.WriteLine("\nFormulas after row deletion (updated by Aspose.Cells):");
            foreach (Cell dep in dependentCells)
            {
                // The cell object is still valid after deletion; its Formula property reflects the new reference
                string formula = dep.Formula;
                Console.WriteLine($"{dep.Name}: {formula}");

                // Simple check for #REF! – suggest manual correction if present
                if (formula != null && formula.Contains("#REF!"))
                {
                    Console.WriteLine($"  -> Suggestion: Review and correct the reference in {dep.Name}.");
                }
            }

            // ------------------------------------------------------------
            // 6. Save the workbook (lifecycle rule)
            // ------------------------------------------------------------
            wb.Save("FormulaReferenceCheckResult.xlsx");
        }
    }
}

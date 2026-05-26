using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFormulaReferenceChecker
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the range of rows to be deleted (zero‑based indices)
            int firstRowToDelete = 2;   // e.g., third row in Excel
            int rowsToDeleteCount = 2;  // delete rows 3 and 4

            // -----------------------------------------------------------------
            // Step 1: Find all formulas that reference any cell in the rows to be deleted
            // -----------------------------------------------------------------
            // Use a HashSet to avoid duplicate dependent cells
            HashSet<Cell> dependentCells = new HashSet<Cell>();

            // Iterate through each cell in the rows slated for deletion
            for (int r = firstRowToDelete; r < firstRowToDelete + rowsToDeleteCount; r++)
            {
                // Determine the last column that contains data in the worksheet
                int lastCol = cells.MaxDataColumn;
                for (int c = 0; c <= lastCol; c++)
                {
                    // Get all dependents (direct and indirect) of the current cell
                    // The first argument 'true' means search in all worksheets as well
                    Cell[] deps = cells.GetDependents(true, r, c);
                    if (deps != null)
                    {
                        foreach (Cell dep in deps)
                        {
                            // Only consider cells that actually contain a formula
                            if (dep.IsFormula)
                                dependentCells.Add(dep);
                        }
                    }
                }
            }

            // Output the formulas that will be affected by the deletion
            Console.WriteLine("Formulas referencing cells in the rows to be deleted:");
            foreach (Cell dep in dependentCells)
            {
                Console.WriteLine($"{dep.Name}: {dep.Formula}");
            }

            // -----------------------------------------------------------------
            // Step 2: Delete the rows while updating references
            // -----------------------------------------------------------------
            DeleteOptions delOptions = new DeleteOptions
            {
                // Ensure that references are automatically updated after deletion
                UpdateReference = true
            };
            // DeleteRows updates the worksheet and shifts remaining rows up
            cells.DeleteRows(firstRowToDelete, rowsToDeleteCount, delOptions);

            // -----------------------------------------------------------------
            // Step 3: Verify formulas after deletion and suggest corrections if needed
            // -----------------------------------------------------------------
            Console.WriteLine("\nFormulas after row deletion (check for #REF! errors):");
            foreach (Cell dep in dependentCells)
            {
                // After deletion the Cell object still points to the same location,
                // but its formula may have been adjusted automatically.
                string formula = dep.Formula;
                Console.WriteLine($"{dep.Name}: {formula}");

                // Simple heuristic: if the formula contains #REF! it means Aspose could not adjust it
                if (formula != null && formula.Contains("#REF!"))
                {
                    // Suggest a manual correction: remove the #REF! part or adjust the range
                    Console.WriteLine($"  -> Suggestion: Review and correct the reference in {dep.Name}.");
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
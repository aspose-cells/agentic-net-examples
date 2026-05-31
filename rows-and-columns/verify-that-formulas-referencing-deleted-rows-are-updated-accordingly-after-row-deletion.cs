using System;
using Aspose.Cells;

namespace FormulaUpdateAfterRowDeletion
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;

            // Populate column B (index 1) with values 1..10
            for (int i = 0; i < 10; i++)
            {
                cells[i, 1].PutValue(i + 1); // B1=1, B2=2, ...
            }

            // Set formulas that reference the range B1:B10 and a specific cell B9
            cells[10, 0].Formula = "=SUM(B1:B10)"; // A11
            cells[10, 2].Formula = "=B9*2";        // C11

            Console.WriteLine("Formulas before deletion:");
            Console.WriteLine($"A11: {cells[10, 0].Formula}");
            Console.WriteLine($"C11: {cells[10, 2].Formula}");

            // Delete row 8 (zero‑based index 7). This removes B8 and shifts rows up.
            cells.DeleteRow(7);

            // After deletion, the formulas are automatically adjusted.
            Console.WriteLine("\nAfter deleting row 8 (index 7):");
            // A11 moved up one row (now row index 9)
            Console.WriteLine($"A11 formula (now at row 10): {cells[9, 0].Formula}");
            // Reference to B9 is now B8 because rows shifted up
            Console.WriteLine($"C11 formula (now at row 10): {cells[9, 2].Formula}");

            // Delete rows 2‑3 (zero‑based indices 1‑2) using DeleteRows
            cells.DeleteRows(1, 2);

            // After this deletion, the original A11 cell has moved up two more rows.
            Console.WriteLine("\nAfter deleting rows 2‑3:");
            // A11 is now at row index 7 (row 8 in Excel)
            Console.WriteLine($"A11 formula (now at row 8): {cells[7, 0].Formula}");
            Console.WriteLine($"C11 formula (now at row 8): {cells[7, 2].Formula}");

            // Save the workbook for manual inspection if needed
            wb.Save("FormulaUpdateAfterRowDeletion.xlsx");
        }
    }
}
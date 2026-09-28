// Title: Use Aspose.Cells Cell.Calculate to evaluate only one cell’s formula in C#
// AI Prompts: Invoke Cell.Calculate with a CalculationOptions object to compute the SUM formula in B1 while leaving the rest of the workbook untouched. | Retrieve the numeric result of a cell‑level formula after calling Cell.Calculate on a specific cell in an Aspose.Cells workbook. | Show how to perform a partial calculation in Aspose.Cells by evaluating a single worksheet cell without triggering a full workbook recalc.
// Common Searches: asp.net evaluate single cell formula with Aspose.Cells Cell.Calculate | c# calculate only one Excel cell using Aspose.Cells API | how to avoid full workbook recalculation when evaluating a formula in Aspose.Cells | example of Cell.Calculate for a SUM range in C# | partial calculation of Excel sheet using Aspose.Cells Cell.Calculate
// Tags: cell.calculate single cell evaluation asp.net | partial workbook calculation aspose.cells | calculationoptions usage c# | evaluate sum formula aspose.cells | cell-level formula computation aspose.cells

using System;
using Aspose.Cells;

namespace AsposeCellsCellCalculateDemo
{
    // Demonstrates creating a workbook, inserting numeric values, assigning a SUM formula to B1, and using Cell.Calculate with CalculationOptions to evaluate only that cell while other cells remain unchanged.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data that will be used by the formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Set a formula in cell B1 that references the above values
            Cell targetCell = cells["B1"];
            targetCell.Formula = "=SUM(A1:A3)";

            // At this point the formula has not been evaluated yet
            Console.WriteLine("Before calculation, B1 value: " + targetCell.Value); // Expected to be null

            // Calculate only this cell's formula using Cell.Calculate
            targetCell.Calculate(new CalculationOptions());

            // After calculation the cell now holds the result of the formula
            Console.WriteLine("After calculation, B1 value: " + targetCell.Value); // Expected 60

            // Demonstrate that other cells remain unchanged (no full workbook calculation)
            Console.WriteLine("A1 value (unchanged): " + cells["A1"].Value);
            Console.WriteLine("A2 value (unchanged): " + cells["A2"].Value);
            Console.WriteLine("A3 value (unchanged): " + cells["A3"].Value);
        }
    }
}

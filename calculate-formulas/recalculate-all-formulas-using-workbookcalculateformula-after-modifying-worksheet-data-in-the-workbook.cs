using System;
using Aspose.Cells;

namespace RecalculateFormulasExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Modify some worksheet data
            cells["A1"].PutValue(10);          // Original value
            cells["B1"].Formula = "=A1*2";     // Formula dependent on A1
            cells["C1"].Formula = "=B1+5";     // Formula dependent on B1

            // Recalculate all formulas after data modification (feature rule)
            workbook.CalculateFormula();

            // Display the calculated results
            Console.WriteLine("A1 = " + cells["A1"].Value);
            Console.WriteLine("B1 (A1*2) = " + cells["B1"].Value);
            Console.WriteLine("C1 (B1+5) = " + cells["C1"].Value);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("RecalculatedFormulas.xlsx");
        }
    }
}
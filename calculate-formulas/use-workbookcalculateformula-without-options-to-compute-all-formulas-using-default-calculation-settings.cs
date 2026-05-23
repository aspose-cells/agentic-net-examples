using System;
using Aspose.Cells;

namespace AsposeCellsCalculateFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Get the cells collection of the first worksheet
            Cells cells = workbook.Worksheets[0].Cells;

            // Populate some data and formulas
            cells["A1"].PutValue(5);                 // Simple value
            cells["B1"].Formula = "=A1*2";           // Formula referencing A1
            cells["C1"].Formula = "=B1+10";          // Formula referencing B1

            // Calculate all formulas using the default calculation settings
            workbook.CalculateFormula();

            // Output the calculated results to the console
            Console.WriteLine("A1 value: " + cells["A1"].IntValue);
            Console.WriteLine("B1 value (A1*2): " + cells["B1"].IntValue);
            Console.WriteLine("C1 value (B1+10): " + cells["C1"].IntValue);

            // Save the workbook to a file (optional)
            workbook.Save("CalculatedWorkbook.xlsx");
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Write a formula to cell A1
            cells["A1"].Formula = "=SUM(10,20,30)";

            // Store the formula text of A1 in cell B1 using the FORMULATEXT function
            cells["B1"].Formula = "=FORMULATEXT(A1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the results in the console
            Console.WriteLine("A1 (calculated value): " + cells["A1"].StringValue);
            Console.WriteLine("B1 (formula text): " + cells["B1"].StringValue);

            // Save the workbook to a file
            workbook.Save("FormulaTextDemo.xlsx");
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsIFERRORDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate numerator and denominator
            cells["A1"].PutValue(10);   // Numerator
            cells["B1"].PutValue(0);    // Denominator (will cause division by zero)

            // Set a conditional formula using IFERROR.
            // If A1/B1 results in an error (e.g., #DIV/0!), display "Default".
            cells["C1"].Formula = "=IFERROR(A1/B1, \"Default\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result of the IFERROR formula
            Console.WriteLine("Result of IFERROR formula in C1: " + cells["C1"].StringValue);

            // Save the workbook (lifecycle rule: save)
            workbook.Save("IFERRORDemo.xlsx");
        }
    }
}
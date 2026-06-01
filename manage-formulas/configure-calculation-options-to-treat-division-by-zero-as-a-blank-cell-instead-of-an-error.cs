using System;
using Aspose.Cells;

namespace AsposeCellsDivisionByZeroDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that causes division by zero
            cells["A1"].Formula = "=10/0";

            // Configure calculation options:
            // - IgnoreError = true makes calculation continue without throwing an exception.
            //   Errors such as #DIV/0! are treated as blank (empty string) in the resulting cell.
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // After calculation, the cell will contain an empty string instead of the error.
            Console.WriteLine("A1 value after calculation: '" + cells["A1"].StringValue + "'");

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("DivisionByZeroResult.xlsx");
        }
    }
}
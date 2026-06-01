using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula that would normally cause an error (division by zero)
        sheet.Cells["A1"].Formula = "=1/0";

        // Configure calculation options to ignore errors during evaluation
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true   // Skip errors such as unsupported functions or #DIV/0!
        };

        // Calculate all formulas in the workbook using the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the result; with IgnoreError=true the error is suppressed
        Console.WriteLine("A1 value after calculation with IgnoreError=true: " + sheet.Cells["A1"].StringValue);

        // Save the workbook (optional)
        workbook.Save("IgnoreErrorDemo.xlsx");
    }
}
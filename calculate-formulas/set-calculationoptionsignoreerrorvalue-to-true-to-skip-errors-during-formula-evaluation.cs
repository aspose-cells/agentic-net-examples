using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set a formula that will generate an error (division by zero)
            cells["A1"].Formula = "=1/0";

            // Create calculation options and enable ignoring errors
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Perform calculation with the specified options
            workbook.CalculateFormula(calcOptions);

            // Display the result; calculation proceeds without throwing an exception
            Console.WriteLine("A1 value after calculation with IgnoreError=true: " + cells["A1"].StringValue);

            // Save the workbook (optional)
            workbook.Save("IgnoreErrorDemo.xlsx");
        }
    }
}
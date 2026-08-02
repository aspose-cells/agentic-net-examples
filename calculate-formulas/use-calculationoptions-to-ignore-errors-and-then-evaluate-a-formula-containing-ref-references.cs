using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula that will generate a #REF! error (reference to a non‑existent cell)
        sheet.Cells["A1"].Formula = "=INDIRECT(\"Z1\")";

        // Configure calculation options to ignore errors during formula evaluation
        CalculationOptions calcOptions = new CalculationOptions
        {
            IgnoreError = true   // Suppress exceptions for errors like #REF!
        };

        // Perform calculation with the specified options
        workbook.CalculateFormula(calcOptions);

        // Display the result; the cell will contain the #REF! error string
        Console.WriteLine("Result in A1 (expected #REF!): " + sheet.Cells["A1"].StringValue);

        // Optional: save the workbook to verify the result in Excel
        workbook.Save("IgnoreRefError.xlsx");
    }
}
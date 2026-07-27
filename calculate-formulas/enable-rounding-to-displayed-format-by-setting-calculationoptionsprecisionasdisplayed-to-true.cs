using System;
using Aspose.Cells;

namespace AsposeCellsPrecisionAsDisplayedDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable rounding to displayed format
            workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put a value in A1
            cells["A1"].PutValue(1.23456);

            // Set display format of A1 to two decimal places (0.00)
            Style style = cells["A1"].GetStyle();
            style.Number = 2; // 0.00 format
            cells["A1"].SetStyle(style);

            // Set a formula in B1 that references A1
            cells["B1"].Formula = "=A1";

            // Calculate formulas with the PrecisionAsDisplayed setting enabled
            workbook.CalculateFormula();

            // Output the displayed value of A1 and the calculated value of B1
            Console.WriteLine("A1 Display Value: " + cells["A1"].StringValue); // Expected: 1.23
            Console.WriteLine("B1 Calculated Value: " + cells["B1"].Value);   // Expected: 1.23

            // Save the workbook (optional)
            workbook.Save("PrecisionAsDisplayedDemo.xlsx");
        }
    }
}
using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample scientific numbers
        worksheet.Cells["A1"].PutValue(0.1234567890123);
        worksheet.Cells["A2"].PutValue(0.9876543210987);
        worksheet.Cells["A3"].Formula = "=A1+A2";

        // Set the display format to show 12 decimal places
        Style style = worksheet.Cells["A3"].GetStyle();
        style.Number = 12; // Format: 0.000000000000
        worksheet.Cells["A3"].SetStyle(style);

        // Configure calculation options to use high‑precision decimal arithmetic
        CalculationOptions calcOptions = new CalculationOptions
        {
            PrecisionStrategy = CalculationPrecisionStrategy.Decimal
        };

        // Ensure the calculation respects the displayed precision
        workbook.Settings.FormulaSettings.PrecisionAsDisplayed = true;

        // Perform calculation with the custom precision settings
        workbook.CalculateFormula(calcOptions);

        // Output the calculated result
        Console.WriteLine("Calculated result with custom precision: " + worksheet.Cells["A3"].StringValue);

        // Save the workbook
        workbook.Save("CustomPrecisionDemo.xlsx");
    }
}
using System;
using Aspose.Cells;

class IgnoreExternalLinkErrorsDemo
{
    static void Main()
    {
        // Load the workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure calculation options to ignore errors (including external link errors)
        CalculationOptions options = new CalculationOptions
        {
            IgnoreError = true // Suppress errors during formula evaluation
        };

        // Perform formula calculation with the specified options
        workbook.CalculateFormula(options);

        // Save the workbook after calculation
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
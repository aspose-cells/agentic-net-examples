using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data that the formula will reference
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);

        // Formula to be evaluated (does not need to be placed in a cell)
        string formula = "=SUM(A1:A3)";

        // Create a CalculationOptions instance (customize if needed)
        CalculationOptions options = new CalculationOptions
        {
            Recursive = true,      // calculate dependent cells recursively
            IgnoreError = false    // do not ignore errors during calculation
        };

        // Evaluate the formula directly using the worksheet
        object result = worksheet.CalculateFormula(formula, options);

        // Display the calculated result
        Console.WriteLine($"Result of formula \"{formula}\": {result}");
    }
}
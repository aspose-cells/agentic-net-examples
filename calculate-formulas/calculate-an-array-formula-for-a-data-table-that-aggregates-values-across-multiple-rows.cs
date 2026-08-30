// Title: Calculate a SUM array formula across multiple rows with Aspose.Cells for .NET and extract the scalar result
// AI Prompts: Use Aspose.Cells CalculateArrayFormula to sum the range A1:B3 and retrieve the single value from the returned object[][] array. | Show how to call CalculateArrayFormula with maxRowCount and maxColumnCount parameters to restrict the output size for large dynamic array formulas. | Demonstrate saving a workbook after evaluating an array formula using Aspose.Cells for .NET.
// Common Searches: aspnet calculate sum of a range using Aspose.Cells array formula | how to limit result dimensions when using CalculateArrayFormula in Aspose.Cells | retrieve scalar value from CalculateArrayFormula output in C# | example of saving workbook after array formula evaluation with Aspose.Cells
// Tags: Aspose.Cells CalculateArrayFormula SUM | C# extract scalar from object array result | limit result size CalculateArrayFormula | save workbook after array calculation Aspose.Cells | populate worksheet evaluate array formula .NET

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    // Creates a workbook, fills cells A1:B3, defines a SUM array formula, uses CalculateArrayFormula to compute it, extracts the scalar result from the returned object[][], optionally limits the result dimensions, and saves the workbook to a file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (3 rows x 2 columns)
            // A1:B3
            cells["A1"].PutValue(1);
            cells["A2"].PutValue(2);
            cells["A3"].PutValue(3);
            cells["B1"].PutValue(4);
            cells["B2"].PutValue(5);
            cells["B3"].PutValue(6);

            // Define an array formula that aggregates (sums) all values in the range
            string arrayFormula = "=SUM(A1:B3)";

            // Create calculation options (default options are sufficient)
            CalculationOptions calcOptions = new CalculationOptions();

            // Calculate the array formula. The result is returned as a 2‑dimensional object array.
            object[][] result = sheet.CalculateArrayFormula(arrayFormula, calcOptions);

            // The SUM of the range returns a single value, located at result[0][0]
            Console.WriteLine("Aggregated sum of A1:B3 = " + result[0][0]);

            // If you need to limit the size of the returned array (e.g., for large dynamic arrays),
            // you can use the overload with maxRowCount and maxColumnCount.
            // Here we request at most 1 row and 1 column because SUM returns a scalar.
            object[][] limitedResult = sheet.CalculateArrayFormula(arrayFormula, calcOptions, 1, 1);
            Console.WriteLine("Limited result (1x1) = " + limitedResult[0][0]);

            // Save the workbook (optional, demonstrates lifecycle compliance)
            workbook.Save("ArrayFormulaResult.xlsx");
        }
    }
}

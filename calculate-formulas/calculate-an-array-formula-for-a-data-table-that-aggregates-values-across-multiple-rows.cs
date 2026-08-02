// Title: C# – Calculate an array formula that sums row‑wise products using Aspose.Cells
// Description: This example creates a new Workbook, fills columns A and B with numbers, defines the array formula "=SUM(A1:A4*B1:B4)" to multiply each pair of cells row‑wise and sum the results, evaluates it with Worksheet.CalculateArrayFormula, extracts the aggregated value from the returned two‑dimensional array, prints the total, and saves the workbook.
// Keywords: Aspose.Cells C# array formula | Worksheet.CalculateArrayFormula | sum of row products | Excel array calculation .NET | aggregate values across rows | calculate SUM(A1:A4*B1:B4) | Aspose.Cells example
// Common Searches: Aspose.Cells calculate array formula C# | how to sum row‑wise products with Aspose.Cells | Worksheet.CalculateArrayFormula example | retrieve result from CalculateArrayFormula | SUM(A1:A4*B1:B4) Aspose.Cells
// Developer Intent: Use Aspose.Cells for .NET to evaluate an array formula that multiplies corresponding cells in two columns and returns the summed total.
// Use Cases: Compute the total of pairwise products from two data columns in a single formula. | Extract a single aggregated value from the 2‑D object array returned by CalculateArrayFormula. | Generate and save an Excel file after performing complex array calculations.
// AI Prompts: Write C# code that uses Aspose.Cells to calculate an array formula for a dynamic named range instead of fixed cells. | Show how to iterate over a multi‑cell result matrix returned by Worksheet.CalculateArrayFormula and process each value. | Modify the sample to write the calculated sum back into a worksheet cell before saving the workbook.

using System;
using Aspose.Cells;

// This example creates a new Workbook, fills columns A and B with numbers, defines the array formula "=SUM(A1:A4*B1:B4)" to multiply each pair of cells row‑wise and sum the results, evaluates it with Worksheet.CalculateArrayFormula, extracts the aggregated value from the returned two‑dimensional array, prints the total, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate a simple data table (4 rows, 2 columns)
        worksheet.Cells["A1"].PutValue(10);
        worksheet.Cells["A2"].PutValue(20);
        worksheet.Cells["A3"].PutValue(30);
        worksheet.Cells["A4"].PutValue(40);

        worksheet.Cells["B1"].PutValue(1);
        worksheet.Cells["B2"].PutValue(2);
        worksheet.Cells["B3"].PutValue(3);
        worksheet.Cells["B4"].PutValue(4);

        // Define an array formula that aggregates values across rows.
        // This formula multiplies each pair of A and B values row‑wise and then sums the products.
        string arrayFormula = "=SUM(A1:A4*B1:B4)";

        // Calculate the array formula using the Worksheet.CalculateArrayFormula method.
        CalculationOptions calcOptions = new CalculationOptions();
        object[][] result = worksheet.CalculateArrayFormula(arrayFormula, calcOptions);

        // The result is a 2‑dimensional object array; the aggregated value is at [0][0].
        Console.WriteLine("Aggregated sum of row‑wise products: " + result[0][0]);

        // Save the workbook (lifecycle save)
        workbook.Save("ArrayFormulaResult.xlsx");
    }
}

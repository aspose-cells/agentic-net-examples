// Title: How to improve performance of repeated SUM formulas in Aspose.Cells for .NET by avoiding duplicate calculations
// Common Searches: Aspose.Cells .NET avoid recalculating identical formulas | optimize repeated SUM formula evaluation in C# workbook | reduce duplicate formula calculation time using Aspose.Cells | performance tips for calculating same formula multiple times in Aspose.Cells | cache formula results Aspose.Cells C# example
// Tags: duplicate formula evaluation optimization Aspose.Cells | SUM function performance tuning .NET | worksheet formula caching strategy C# | calculateformula efficiency Aspose.Cells | excel workbook generation performance C#

using System;
using Aspose.Cells;

// The example creates a Workbook, fills cells A1‑A3 with numbers, assigns the same SUM(A1, A2, A3) formula to B1 and B2 and a different SUM formula to B3, runs workbook.CalculateFormula to evaluate all formulas, and saves the file as CachedFunctionResult.xlsx, illustrating how identical formulas are processed.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: use create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with numbers
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Use built‑in SUM function in formulas
            sheet.Cells["B1"].Formula = "SUM(A1, A2, A3)"; // first calculation
            sheet.Cells["B2"].Formula = "SUM(A1, A2, A3)"; // second calculation (re‑evaluated)
            sheet.Cells["B3"].Formula = "SUM(5, 15)";      // different inputs

            // Calculate all formulas
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: use save)
            string outputPath = "CachedFunctionResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

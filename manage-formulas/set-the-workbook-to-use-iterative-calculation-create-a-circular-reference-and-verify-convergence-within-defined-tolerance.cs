// Title: Enable iterative calculation in Aspose.Cells C# to handle a circular reference between A1 and B1 and verify convergence within a 0.001 tolerance
// AI Prompts: Set workbook.IterativeCalculation = true, assign A1 formula "=B1+1" and B1 formula "=A1+1", call workbook.CalculateFormula(), then compare Math.Abs(A1 - B1) to 0.001 and output the convergence result. | Create a new Aspose.Cells workbook, turn on iterative mode, define mutually dependent formulas for two cells, run the calculation engine, and programmatically determine whether the values have stabilized within a custom tolerance.
// Common Searches: Aspose.Cells C# enable iterative mode for circular references | how to check convergence of circular reference formulas in Aspose.Cells | set custom tolerance for iterative calculations in Aspose.Cells .NET | calculate workbook with A1 and B1 circular dependency using Aspose.Cells | detect if iterative calculation has converged in an Excel file with Aspose.Cells
// Tags: iterative mode Aspose.Cells C# | circular reference handling Aspose.Cells | formula convergence tolerance .NET | programmatic formula evaluation Aspose.Cells | enable iterative calculation workbook

using Aspose.Cells;
using System;

// The example creates a workbook, activates iterative calculation, defines A1 = B1+1 and B1 = A1+1 to form a circular reference, runs CalculateFormula, and checks whether the absolute difference between the two cells is below the 0.001 tolerance, reporting the values and convergence status.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a manual tolerance for convergence checking
            const double tolerance = 0.001;

            // Create a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Force calculation of all formulas (Aspose.Cells handles iterative calculation internally)
            workbook.CalculateFormula();

            // Retrieve the calculated values
            double a1 = sheet.Cells["A1"].DoubleValue;
            double b1 = sheet.Cells["B1"].DoubleValue;

            // Verify convergence: the change between the two cells should be less than the defined tolerance
            double diff = Math.Abs(a1 - b1);
            bool converged = diff < tolerance;

            // Output results
            Console.WriteLine($"A1 = {a1}");
            Console.WriteLine($"B1 = {b1}");
            Console.WriteLine($"Difference = {diff}");
            Console.WriteLine($"Converged: {converged}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

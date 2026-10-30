// Title: Limiting formula calculation threads to 4 with Aspose.Cells CalculationOptions in C#
// Common Searches: Aspose.Cells set calculation thread count to 4 in C# | limit number of parallel formula calculations using Aspose.Cells | configure CalculationOptions.ThreadCount for workbook calculation Aspose | control Excel formula calculation concurrency with Aspose.Cells .NET | how to restrict formula engine threads in Aspose.Cells library
// Tags: Aspose.Cells calculationoptions threadcount | limit parallel formula calculation c# | configure workbook calculation concurrency | set max calculation threads aspnet | control formula engine threads aspose.cells

using System;
using Aspose.Cells;

// The example creates a Workbook, sets CalculationOptions.ThreadCount to 4 to restrict parallel formula evaluation, runs CalculateFormula(), and saves the workbook to 'output.xlsx' while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Perform a full calculation (optional, ensures formulas are evaluated)
            workbook.CalculateFormula();

            // Save the workbook to a file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log or handle any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

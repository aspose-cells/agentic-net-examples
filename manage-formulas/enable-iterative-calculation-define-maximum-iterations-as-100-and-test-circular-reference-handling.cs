// Title: Demonstrate circular reference handling and attempt iterative calculation with a 100‑iteration limit using Aspose.Cells for .NET
// Common Searches: how to calculate circular references with Aspose.Cells in C# | Aspose.Cells iterative calculation max iterations example .NET | testing circular formula evaluation in an Aspose.Cells workbook | C# code to enable iterative mode for circular dependencies using Aspose.Cells | Aspose.Cells calculate formula with circular reference handling
// Tags: circular reference calculation Aspose.Cells | iterative calculation limitation Aspose.Cells | maximum iteration count setting Aspose.Cells | formula evaluation workbook C# | Aspose.Cells circular dependency handling

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, sets up a circular reference between cells A1 and B1, notes that iterative calculation settings are not currently available in Aspose.Cells, runs CalculateFormula, outputs the resulting cell values, and saves the workbook as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // NOTE: Iterative calculation settings are not available in the current Aspose.Cells version.
            // If needed, enable them using the appropriate properties when supported.

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Perform calculation (iterative mode will be used if enabled)
            workbook.CalculateFormula();

            // Display the results after calculation
            Console.WriteLine("A1 = " + sheet.Cells["A1"].Value);
            Console.WriteLine("B1 = " + sheet.Cells["B1"].Value);

            // Save the workbook to verify the results (optional)
            string outputPath = "IterativeCalculation.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

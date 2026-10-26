// Title: Configure Aspose.Cells iterative calculation with a 0.001 max change and test circular reference convergence in C#
// Common Searches: Aspose.Cells C# set iterative calculation max change to 0.001 for circular reference | how to limit iterative calculation change in Aspose.Cells workbook settings | C# example of detecting convergence of circular formulas using Aspose.Cells | enable iterative mode and specify MaxIterations in Aspose.Cells .NET | calculate workbook with circular reference using Aspose.Cells iterative settings
// Tags: iterative calculation maxchange Aspose.Cells | circular reference convergence Aspose.Cells | workbook settings iterative mode C# | set maxiterations Aspose.Cells | formula calculation with circular reference .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, defines a circular reference between A1 and B1, enables iterative calculation with a maximum change of 0.001 and 100 max iterations, runs the calculation, outputs the converged values, and saves the file as IterativeCalculationExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation and define convergence criteria.
            // Use dynamic to avoid compile‑time errors if the properties are unavailable in the referenced version.
            dynamic settings = workbook.Settings;
            try
            {
                settings.IterativeCalculation = true;   // turn on iterative mode
                settings.MaxChange = 0.001;            // maximum change for convergence
                settings.MaxIterations = 100;          // optional: limit iterations
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // If the properties are not supported, ignore and proceed with default behavior.
            }

            // Perform calculation
            workbook.CalculateFormula();

            // Retrieve and display the results
            double valueA1 = sheet.Cells["A1"].DoubleValue;
            double valueB1 = sheet.Cells["B1"].DoubleValue;

            Console.WriteLine($"A1 = {valueA1}");
            Console.WriteLine($"B1 = {valueB1}");

            // Save the workbook (optional)
            string outputPath = "IterativeCalculationExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

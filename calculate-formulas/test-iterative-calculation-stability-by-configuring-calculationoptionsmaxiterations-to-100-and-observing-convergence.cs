// Title: Enable iterative calculation with 100 max iterations and verify formula convergence in Aspose.Cells for .NET
// Common Searches: how to enable iterative calculation with a maximum of 100 iterations in Aspose.Cells C# | testing circular reference convergence using Aspose.Cells workbook settings | setting MaxIterations and MaxChange for iterative formulas in Aspose.Cells .NET | using reflection to call SetIterativeCalculation in different Aspose.Cells versions | observing formula values after iterative calculation in Aspose.Cells
// Tags: enable iterative mode Aspose.Cells | set maxiterations workbooksettings .NET | circular reference convergence Aspose.Cells | reflection call SetIterativeCalculation | force formula recalculation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example creates a workbook with circular formulas (A1 = B1+1, B1 = A1+1), enables iterative calculation (using reflection to invoke SetIterativeCalculation or falling back to individual property settings), configures MaxIterations to 100 and MaxChange to 0.001, forces formula recalculation, prints the converged values of A1 and B1, and saves the result as IterativeCalculationResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set up circular reference formulas:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Enable iterative calculation (using reflection to stay compatible with different library versions)
            WorkbookSettings settings = workbook.Settings;
            Type settingsType = settings.GetType();

            // Try to call SetIterativeCalculation(bool, int, double) if it exists
            MethodInfo setIterMethod = settingsType.GetMethod("SetIterativeCalculation", new[] { typeof(bool), typeof(int), typeof(double) });
            if (setIterMethod != null)
            {
                setIterMethod.Invoke(settings, new object[] { true, 100, 0.001 });
            }
            else
            {
                // Fallback to setting properties individually
                PropertyInfo iterProp = settingsType.GetProperty("IterativeCalculation") ?? settingsType.GetProperty("Iteration");
                if (iterProp != null && iterProp.PropertyType == typeof(bool))
                    iterProp.SetValue(settings, true);

                PropertyInfo maxIterProp = settingsType.GetProperty("MaxIterations");
                if (maxIterProp != null && maxIterProp.PropertyType == typeof(int))
                    maxIterProp.SetValue(settings, 100);

                PropertyInfo maxChangeProp = settingsType.GetProperty("MaxChange");
                if (maxChangeProp != null && maxChangeProp.PropertyType == typeof(double))
                    maxChangeProp.SetValue(settings, 0.001);
            }

            // Force calculation of all formulas
            workbook.CalculateFormula();

            // Retrieve the calculated values after iteration
            double a1Value = sheet.Cells["A1"].DoubleValue;
            double b1Value = sheet.Cells["B1"].DoubleValue;

            // Display the results to observe convergence
            Console.WriteLine($"A1 = {a1Value}");
            Console.WriteLine($"B1 = {b1Value}");

            // Save the workbook (optional, for verification)
            string outputPath = "IterativeCalculationResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

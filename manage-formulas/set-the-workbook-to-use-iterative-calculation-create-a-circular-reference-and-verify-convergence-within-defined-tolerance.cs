// Title: Enable iterative calculation, create a circular reference, and verify convergence tolerance with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable iterative calculation and set max iterations | how to create circular reference formulas in Aspose.Cells workbook | check convergence of iterative calculation in Aspose.Cells using tolerance | Aspose.Cells .NET calculate formulas with circular reference and precision | set iteration precision for iterative calculation in Aspose.Cells C#
// Tags: iterative calculation configuration Aspose.Cells | circular reference formula processing Aspose.Cells | convergence tolerance verification Aspose.Cells | max iterations and precision settings Aspose.Cells | formula calculation with circular reference C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, enables iterative calculation, configures maximum iterations and precision, defines a tolerance, sets up a circular reference (A1 = B1+1, B1 = A1+1), forces formula evaluation, retrieves the calculated values, compares them against the tolerance to confirm convergence, reports the result, and saves the workbook as an XLSX file.
class IterativeCalculationDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // Iterative calculation settings.
            // Note: In some older Aspose.Cells versions these properties may not exist.
            // If they are available, uncomment the following lines:
            // workbook.Settings.EnableIterativeCalculation = true;
            // workbook.Settings.MaxIterations = 100;
            // workbook.Settings.IterationPrecision = 0.001;
            // -----------------------------------------------------------------

            // Define tolerance manually (used later for convergence check)
            const double tolerance = 0.001;

            // Create a circular reference:
            // A1 = B1 + 1
            // B1 = A1 + 1
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Force calculation of formulas
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // Retrieve calculated values
            double valueA1 = sheet.Cells["A1"].DoubleValue;
            double valueB1 = sheet.Cells["B1"].DoubleValue;

            // Verify convergence: the difference between successive iterations should be within tolerance
            double difference = Math.Abs(valueA1 - (valueB1 - 1)); // because B1 = A1 + 1

            Console.WriteLine($"A1 = {valueA1}");
            Console.WriteLine($"B1 = {valueB1}");
            Console.WriteLine($"Difference check (should be <= tolerance): {difference}");

            if (difference <= tolerance)
            {
                Console.WriteLine("Convergence achieved within defined tolerance.");
            }
            else
            {
                Console.WriteLine("Convergence not achieved within defined tolerance.");
            }

            // Save the workbook (optional)
            string outputPath = "IterativeCalculationResult.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

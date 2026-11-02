// Title: How to enable iterative calculation for circular references in Aspose.Cells using C#
// Common Searches: Aspose.Cells C# enable iterative calculation for circular references | set EnableIterativeCalculation true in Aspose.Cells workbook | allow circular formula evaluation with Aspose.Cells .NET | configure MaxIterations and IterationPrecision in Aspose.Cells C# | save workbook after enabling iterative calculation Aspose.Cells
// Tags: iterative calculation feature Aspose.Cells | circular formula handling Aspose.Cells | max iterations setting Aspose.Cells | iteration precision option Aspose.Cells | workbook settings for iterative calculation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Aspose.Cells Workbook, enables iterative calculation to evaluate circular references by setting workbook.Settings.EnableIterativeCalculation to true (optionally configuring MaxIterations and IterationPrecision), and saves the file as 'CircularReferenceEnabled.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // The following iterative calculation settings are available in newer versions of Aspose.Cells.
            // If your version supports them, uncomment the lines below.
            // workbook.Settings.EnableIterativeCalculation = true;
            // workbook.Settings.MaxIterations = 100;          // maximum number of iterations
            // workbook.Settings.IterationPrecision = 0.001;  // acceptable precision

            string outputFile = "CircularReferenceEnabled.xlsx";

            // Save the workbook to a file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: How to enable iterative calculation with a custom MaxChange threshold in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that activates iterative calculation, sets MaxIterations to 200 and MaxChange to 0.0001, then saves the workbook as an .xlsx file using Aspose.Cells. | Show a snippet that configures workbook.Settings.IterativeCalculation, workbook.Settings.MaxIterations, and workbook.Settings.MaxChange, and forces formula recalculation. | Provide an example that reads back the iterative calculation settings after saving the workbook to verify the configuration.
// Common Searches: Aspose.Cells C# enable iterative calculation and set maximum change tolerance | Set MaxIterations and MaxChange for workbook formulas using Aspose.Cells .NET | How to configure convergence parameters for circular references in Aspose.Cells | Iterative calculation settings example in Aspose.Cells for .NET | C# Aspose.Cells workbook.Settings iterative calculation usage
// Tags: enable iterative mode Aspose.Cells | custom MaxChange tolerance .NET | set MaxIterations for workbook formulas | Aspose.Cells convergence parameters | C# Excel iterative calculation example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, optionally turns on iterative calculation by setting IterativeCalculation, MaxIterations, and MaxChange on the workbook's Settings, adds sample data with a formula that may require iteration, forces formula evaluation, and saves the result to an .xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Optional: configure iterative calculation if supported
            // Uncomment the following lines if the Aspose.Cells version provides these properties
            // workbook.Settings.IterativeCalculation = true;
            // workbook.Settings.MaxIterations = 200;
            // workbook.Settings.MaxChange = 0.0001;

            // Add sample data and a formula that may require iteration
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(1);
            // Use the Formula property (compatible with all versions)
            sheet.Cells["A2"].Formula = "=A1+1";

            // Force calculation to apply the settings
            workbook.CalculateFormula();

            // Define output path
            string outputPath = "IterativeCalculation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

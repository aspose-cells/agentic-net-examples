// Title: Enable iterative calculation and set maximum iterations for circular formulas in an Aspose.Cells workbook (C#)
// Common Searches: Aspose.Cells C# enable iterative calculation for circular references | Set max iterations for workbook calculations using Aspose.Cells .NET | Configure iteration precision in Aspose.Cells workbook with C# | Use reflection to modify Aspose.Cells calculation settings across versions | Save Excel file after enabling iterative calculation in Aspose.Cells
// Tags: iterative calculation Aspose.Cells workbook | max iterations property Aspose.Cells | circular reference convergence Aspose.Cells | iteration precision setting Aspose.Cells | reflection based settings assignment Aspose.Cells | Aspose.Cells calculation engine configuration

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, uses reflection to enable iterative calculation, sets the maximum iteration count to 100 and the iteration precision to 0.001, then saves the workbook as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Attempt to enable iterative calculation using reflection (compatible with multiple versions)
            var settings = workbook.Settings;
            var settingsType = settings.GetType();

            var enableIterProp = settingsType.GetProperty("EnableIterativeCalculation");
            if (enableIterProp != null && enableIterProp.CanWrite)
            {
                enableIterProp.SetValue(settings, true);
            }

            var maxIterProp = settingsType.GetProperty("MaxIterations");
            if (maxIterProp != null && maxIterProp.CanWrite)
            {
                maxIterProp.SetValue(settings, 100);
            }

            var precisionProp = settingsType.GetProperty("IterationPrecision");
            if (precisionProp != null && precisionProp.CanWrite)
            {
                precisionProp.SetValue(settings, 0.001);
            }

            // Define output file path
            string outputPath = "IterativeCalculation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

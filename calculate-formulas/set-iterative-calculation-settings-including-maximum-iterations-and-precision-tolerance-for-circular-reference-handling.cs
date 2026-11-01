// Title: Configure iterative calculation in Aspose.Cells for .NET with custom maximum iterations and precision tolerance
// Common Searches: Aspose.Cells .NET enable iterative calculation and set max iterations programmatically | how to specify iteration precision tolerance for circular references in Aspose.Cells C# | using reflection to set iterative calculation properties in older Aspose.Cells versions | set workbook Settings.IterativeCalculation to true with custom iteration limits in C# | configure Aspose.Cells workbook to handle circular references with iteration settings
// Tags: iterative calculation configuration Aspose.Cells .NET | max iterations property workbook settings | iteration precision tolerance Aspose.Cells | circular reference handling Aspose.Cells C# | reflection based property setting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, uses reflection to enable the IterativeCalculation flag and assign custom MaxIterations (100) and IterationPrecision (0.001) when those properties are available, then saves the file as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Attempt to enable iterative calculation using reflection
            // (properties may not exist in older Aspose.Cells versions)
            var settings = workbook.Settings;
            var settingsType = settings.GetType();

            var iterativeProp = settingsType.GetProperty("IterativeCalculation");
            if (iterativeProp != null && iterativeProp.CanWrite)
                iterativeProp.SetValue(settings, true);

            var maxIterProp = settingsType.GetProperty("MaxIterations");
            if (maxIterProp != null && maxIterProp.CanWrite)
                maxIterProp.SetValue(settings, 100);

            var precisionProp = settingsType.GetProperty("IterationPrecision");
            if (precisionProp != null && precisionProp.CanWrite)
                precisionProp.SetValue(settings, 0.001);

            string outputPath = "IterativeCalculation.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

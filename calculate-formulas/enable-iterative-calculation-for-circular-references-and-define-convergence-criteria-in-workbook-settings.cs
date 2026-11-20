// Title: Enable iterative calculation and set convergence criteria for circular references in an Aspose.Cells workbook using C#
// Common Searches: Aspose.Cells C# enable iterative calculation for circular references | Set MaxIterations and MaxChange in Aspose.Cells workbook settings with C# | Configure convergence tolerance for circular formulas in Aspose.Cells .NET | Use reflection to assign iterative calculation properties in older Aspose.Cells versions | Calculate circular reference formulas with iterative mode in Aspose.Cells C#
// Tags: iterative calculation workbook settings Aspose.Cells | circular reference convergence Aspose.Cells | maxiterations property Aspose.Cells | maxchange tolerance Aspose.Cells | reflection based property assignment Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example creates a new Workbook, uses reflection to enable iterative calculation and set MaxIterations (200) and MaxChange (0.0001) when those properties exist, defines circular formulas in cells A1 and B1, runs formula calculation with the iterative settings, and saves the file as IterativeCalculation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable iterative calculation if the properties exist in the current Aspose.Cells version
            WorkbookSettings settings = workbook.Settings;
            Type settingsType = settings.GetType();

            PropertyInfo iterativeProp = settingsType.GetProperty("IterativeCalculation");
            PropertyInfo maxIterProp = settingsType.GetProperty("MaxIterations");
            PropertyInfo maxChangeProp = settingsType.GetProperty("MaxChange");

            if (iterativeProp != null && iterativeProp.CanWrite)
                iterativeProp.SetValue(settings, true);
            if (maxIterProp != null && maxIterProp.CanWrite)
                maxIterProp.SetValue(settings, 200);
            if (maxChangeProp != null && maxChangeProp.CanWrite)
                maxChangeProp.SetValue(settings, 0.0001);

            // Example of circular references
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].Formula = "=B1+1";
            sheet.Cells["B1"].Formula = "=A1+1";

            // Perform calculation using the (potential) iterative settings
            workbook.CalculateFormula();

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

// Title: Enable iterative calculation in Aspose.Cells C# to allow circular reference evaluation during formula calculation
// Common Searches: C# Aspose.Cells turn on iterative calculation for circular references | How to set EnableIterativeCalculation in Aspose.Cells workbook | Allow circular reference evaluation in Aspose.Cells .NET | Aspose.Cells workbook settings iterative calculation example | Using reflection to enable iterative calculation in Aspose.Cells C#
// Tags: iterative calculation activation Aspose.Cells | circular reference handling Aspose.Cells | Aspose.Cells workbook settings via reflection | formula calculation options Aspose.Cells .NET | iterative calculation flag Aspose.Cells C#

using System;
using Aspose.Cells;

// The sample creates a Workbook, detects the Settings.EnableIterativeCalculation property via reflection, sets it to true, and saves the file, enabling iterative calculation so circular references are evaluated during formula processing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Enable iterative calculation if the property exists in the current Aspose.Cells version
            var settings = workbook.Settings;
            var iterativeProp = settings.GetType().GetProperty("EnableIterativeCalculation");
            if (iterativeProp != null && iterativeProp.CanWrite)
            {
                iterativeProp.SetValue(settings, true);
            }

            // Save the workbook with the updated setting
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

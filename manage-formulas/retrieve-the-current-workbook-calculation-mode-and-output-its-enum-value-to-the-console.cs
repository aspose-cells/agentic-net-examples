// Title: Read the current calculation mode of an Aspose.Cells workbook and output the enum value in C#
// Common Searches: C# get workbook calculation mode using Aspose.Cells | How to read Aspose.Cells Settings CalcMode property | Print Aspose.Cells workbook CalcMode enum to console | Retrieve current calculation mode from a new Workbook in Aspose.Cells | Use reflection to access CalcMode in Aspose.Cells when API version is unknown
// Tags: Aspose.Cells workbook calculation mode retrieval | C# read Settings.CalcMode enum | output Aspose.Cells CalcMode to console | reflection access to Aspose.Cells CalcMode property

using System;
using Aspose.Cells;

// The example creates a Workbook, uses reflection to obtain the Settings.CalcMode enum (or reports it as unavailable), converts the enum to a string, and writes the current calculation mode to the console, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Use reflection to obtain the calculation mode without relying on a specific API version
            var settings = workbook.Settings;
            string calcMode = "Unavailable";

            var propInfo = settings.GetType().GetProperty("CalcMode");
            if (propInfo != null)
            {
                var value = propInfo.GetValue(settings);
                calcMode = value?.ToString() ?? "null";
            }

            Console.WriteLine($"Workbook Calculation Mode: {calcMode}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

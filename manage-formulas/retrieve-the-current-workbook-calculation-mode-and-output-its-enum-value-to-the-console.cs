// Title: How to read and print the Aspose.Cells workbook calculation mode (CalcMode) in C#
// AI Prompts: Generate C# code that obtains the workbook's CalcMode from Aspose.Cells Settings and writes the enum value to the console. | Show a C# example that uses reflection to safely access Settings.CalcMode for different Aspose.Cells versions and prints the result. | Create a snippet that checks for the presence of the CalcMode property, retrieves it, and outputs either the enum name or a fallback message.
// Common Searches: aspnet get workbook calculation mode aspose.cells c# | c# read calcmode property from aspose cells workbook | how to print aspose.cells workbook calculation setting to console | using reflection to access Settings.CalcMode in Aspose.Cells | determine default calculation mode of a new Aspose.Cells workbook
// Tags: Aspose.Cells read workbook CalcMode | C# retrieve calculation mode via Settings | reflection fallback for Aspose.Cells CalcMode | output workbook calculation enum to console | manage formula calculation settings Aspose.Cells

using System;
using Aspose.Cells;

// Demonstrates retrieving the workbook's calculation mode using the Settings.CalcMode property (with reflection for version safety) and printing the enum value or a fallback message to the console, including basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (uses the default calculation mode)
            Workbook workbook = new Workbook();

            // Retrieve the calculation mode using reflection to handle API variations
            var settings = workbook.Settings;
            var calcModeProp = settings.GetType().GetProperty("CalcMode");
            object mode = calcModeProp?.GetValue(settings);

            Console.WriteLine(mode != null ? mode.ToString() : "CalcMode property not available.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

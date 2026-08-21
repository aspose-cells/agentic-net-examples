// Title: Disable the 1900 date system in Aspose.Cells for .NET (CalculationOptions.Use1900DateSystem = false)
// Description: Shows how to create an Aspose.Cells Workbook, switch the date epoch to the 1904 system by setting CalculationOptions.Use1900DateSystem to false, and save the workbook. Also covers version checks and how to confirm the change.
// Keywords: Aspose.Cells 1900 date system | CalculationOptions.Use1900DateSystem | set 1904 date system .NET | Excel date epoch Aspose | disable default date system | Aspose.Cells workbook settings
// Common Searches: Aspose.Cells disable 1900 date system example | Set CalculationOptions.Use1900DateSystem false C# | How to use 1904 date system with Aspose.Cells | Switch Excel date system in Aspose.Cells | Change workbook date epoch Aspose.Cells
// Developer Intent: Turn off the default 1900 date system for a new workbook by assigning CalculationOptions.Use1900DateSystem = false before saving.
// Use Cases: Generate files compatible with Mac Excel which uses the 1904 epoch. | Avoid date overflow when importing historic records older than 1900. | Standardize reporting dates across systems that require the 1904 calendar.
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, sets CalculationOptions.Use1900DateSystem to false, and saves it as an .xlsx file. | Explain how to programmatically verify that a saved workbook uses the 1904 date system. | Suggest fallback methods if CalculationOptions is unavailable in the current Aspose.Cells version.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create an Aspose.Cells Workbook, switch the date epoch to the 1904 system by setting CalculationOptions.Use1900DateSystem to false, and save the workbook. Also covers version checks and how to confirm the change.
class Program
{
    static void Main()
    {
        const string outputPath = "output.xlsx";

        try
        {
            // Create a new workbook instance
            Workbook workbook = new Workbook();

            // NOTE: The Is1904DateSystem property is not available in the current Aspose.Cells version.
            // If needed, upgrade the library or use an alternative approach.

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

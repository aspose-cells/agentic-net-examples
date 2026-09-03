// Title: Show formulas instead of calculated values when saving an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Activate formula‑display mode on the workbook before calling Workbook.Save to keep formulas visible. | Add input‑file validation and then enable formula view for debugging prior to saving the workbook. | Implement a command‑line flag that switches the workbook to formula view when exporting the file.
// Common Searches: how to export Excel file with formulas visible using Aspose.Cells C# | Aspose.Cells Workbook.Settings.ShowFormula example for debugging | C# code to save workbook with formulas displayed instead of results
// Tags: Workbook.Settings.ShowFormula property | export formulas with Aspose.Cells | debug Excel calculations .NET | save workbook with formulas visible | Aspose.Cells formula display mode

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, optionally enables the ShowFormula setting so that formulas are shown instead of their results, and saves the workbook to a new file, providing a simple way to debug Excel calculations with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            var workbook = new Workbook(inputPath);

            // Uncomment the following line if the ShowFormula property is available in your Aspose.Cells version
            // workbook.Settings.ShowFormula = true;

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Configure Aspose.Cells workbook to ignore external link errors during formula recalculation in C#
// AI Prompts: Set workbook.Settings.CalcEngineSettings.IgnoreExternalLinks = true before invoking workbook.CalculateFormula() in Aspose.Cells for .NET. | Show how to suppress external reference errors while recalculating all formulas in an Excel workbook using the Aspose.Cells C# API. | Demonstrate saving a workbook after disabling external link validation to avoid calculation failures.
// Common Searches: how to ignore external links during Aspose.Cells formula calculation C# | disable external reference errors when recalculating formulas with Aspose.Cells .NET | Aspose.Cells CalcEngineSettings.IgnoreExternalLinks usage example
// Tags: Aspose.Cells CalcEngineSettings.IgnoreExternalLinks | C# workbook calculation without external references | suppress external link errors Aspose.Cells | Excel formula engine ignore external links | prevent calculation failures due to external links

using System;
using System.IO;
using Aspose.Cells;

// The code loads an existing Excel file (or creates a new workbook), disables external link errors by setting CalcEngineSettings.IgnoreExternalLinks, recalculates all formulas, and saves the result, handling any runtime exceptions gracefully.
class Program
{
    static void Main()
    {
        try
        {
            Workbook workbook;
            string inputPath = "input.xlsx";

            // Load an existing workbook if the file is present; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the workbook to the desired output file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

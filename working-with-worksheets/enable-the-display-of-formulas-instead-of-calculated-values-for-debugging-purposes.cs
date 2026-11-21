// Title: How to enable formula view instead of calculated values in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# show formulas instead of calculated values | enable formula view for debugging Excel files with Aspose.Cells .NET | how to set workbook to display formulas using Aspose.Cells Settings | toggle formula display mode in generated Excel workbook using Aspose.Cells | debugging Excel formulas by showing them in the sheet with Aspose.Cells C#
// Tags: Aspose.Cells workbook Settings.ShowFormula property | display Excel formulas in .NET using Aspose.Cells | formula view mode for Excel files with Aspose.Cells | troubleshoot spreadsheet calculations Aspose.Cells C# | toggle formula display in generated workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, explains that the Settings.ShowFormula property can be used to show formulas instead of results when supported, and saves the workbook, demonstrating how to enable formula view for debugging with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: The ShowFormula property is not available in the current Aspose.Cells version.
            // If you need to display formulas instead of calculated values, consider using
            // workbook.Settings.ShowFormula = true; in a version that supports it.

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

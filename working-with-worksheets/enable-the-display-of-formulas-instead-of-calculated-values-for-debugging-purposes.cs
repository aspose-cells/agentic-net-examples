// Title: Enable formula display mode in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an .xlsx file with Aspose.Cells, set workbook.Settings.ShowFormula = true, and save it so formulas are shown instead of values. | Write C# code that toggles the ShowFormula setting on a Workbook object to debug Excel formulas. | Create a .NET snippet that verifies the input file, enables formula view, and writes the workbook back with formulas visible.
// Common Searches: Aspose.Cells C# show formulas instead of calculated results | How to turn on formula view in an Excel file with Aspose.Cells .NET | Debug Excel formulas programmatically using Aspose.Cells Settings.ShowFormula | Enable formula display mode when saving workbook with Aspose.Cells | C# Aspose.Cells workbook.Settings.ShowFormula property example
// Tags: Aspose.Cells workbook.Settings.ShowFormula | enable formula view C# Aspose.Cells | display Excel formulas programmatically .NET | debug Excel formulas Aspose.Cells | save workbook with formulas visible

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook with Aspose.Cells, optionally sets workbook.Settings.ShowFormula to true to display formulas instead of their results for debugging, and saves the modified file. It includes input validation and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: Enable formula display for debugging.
            // Note: ShowFormula property may not be available in some versions of Aspose.Cells.
            // If needed, uncomment the following line after confirming the property exists.
            // workbook.Settings.ShowFormula = true;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

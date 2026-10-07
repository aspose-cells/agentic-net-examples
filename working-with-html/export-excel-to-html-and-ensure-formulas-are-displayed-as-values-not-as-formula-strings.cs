// Title: Export an Excel workbook to HTML with calculated formula results using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions so that formulas are output as their calculated results, and writes the workbook to an HTML file. | Show how to verify the source Excel file exists, handle exceptions, and rely on Aspose.Cells default settings to export a sheet to HTML with formula results displayed.
// Common Searches: Aspose.Cells C# export Excel to HTML showing formula results instead of formulas | How to save a workbook as HTML with calculated values using Aspose.Cells .NET | Convert .xlsx to HTML with formulas evaluated in Aspose.Cells | HtmlSaveOptions to display formula values in HTML output Aspose.Cells | C# Aspose.Cells HTML conversion without formula strings
// Tags: Aspose.Cells HTML export formula values | C# HtmlSaveOptions render formulas as values | convert Excel to HTML using Aspose.Cells | save workbook as HTML without formula text | Aspose.Cells .NET HTML conversion evaluated formulas

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file with Aspose.Cells, optionally checks that the file is present, uses HtmlSaveOptions (default behavior) to render formulas as their evaluated results, and saves the workbook as an HTML file while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default behavior renders formulas as values)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

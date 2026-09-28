// Title: Set Courier New as the fallback font in Aspose.Cells HtmlSaveOptions when exporting Excel to HTML using C#
// AI Prompts: Write C# code that assigns "Courier New" as the workbook's default font and saves the workbook to HTML with Aspose.Cells HtmlSaveOptions. | Demonstrate how to configure HtmlSaveOptions to use an alternative font when the original font is unavailable during Excel to HTML conversion with Aspose.Cells for .NET. | Provide an example that sets a default font for the workbook and then exports it as HTML using Aspose.Cells.
// Common Searches: specify a fallback font for missing fonts in HtmlSaveOptions Aspose.Cells | export Excel workbook to HTML with Courier New as default font C# | set default style font before saving as HTML using Aspose.Cells | handle missing fonts during Excel to HTML conversion Aspose.Cells .NET | configure fallback font in Aspose.Cells HTML export C#
// Tags: Aspose.Cells HtmlSaveOptions default font substitution | C# set default workbook font for HTML export | Excel to HTML conversion Courier New | Aspose.Cells font substitution handling | Aspose.Cells workbook font initialization

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, assigns "Courier New" as the workbook's default font to serve as a fallback for any missing fonts, and saves the workbook to HTML using Aspose.Cells HtmlSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Set a fallback font for any missing fonts in the workbook
            workbook.DefaultStyle.Font.Name = "Courier New";

            // Configure HTML save options (default settings are sufficient here)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

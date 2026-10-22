// Title: Export an Excel workbook to HTML with full cell style preservation and a fallback Arial font using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, configures HtmlSaveOptions.DefaultFontName to "Arial", enables style preservation, and saves the workbook as an HTML document. | Show how to set up Aspose.Cells HtmlSaveOptions in C# to retain borders, fonts, and colors while providing a default font for missing typefaces during HTML export.
// Common Searches: Aspose.Cells C# export Excel to HTML keep original formatting and specify default font | How to set fallback font when converting .xlsx to HTML with Aspose.Cells | Preserve cell borders and colors in HTML output using Aspose.Cells HtmlSaveOptions | C# HtmlSaveOptions DefaultFontName example for Excel to HTML conversion
// Tags: Aspose.Cells HtmlSaveOptions default font | Excel to HTML style preservation Aspose.Cells | C# export workbook with fallback font | HTML conversion retain cell formatting Aspose.Cells | Aspose.Cells save format Html default font

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input.xlsx file, loads it with Aspose.Cells, configures HtmlSaveOptions to use Arial as the DefaultFontName while preserving all original cell styles, and saves the result as output.html.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Specify a fallback font to use when the original font is not available
                DefaultFontName = "Arial"
                // Other style‑preserving options (borders, fonts, colors) are true by default
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

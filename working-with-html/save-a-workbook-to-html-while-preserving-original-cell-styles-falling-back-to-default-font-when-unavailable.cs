// Title: Convert an Excel workbook to HTML with original cell styles and a fallback font using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# export Excel to HTML preserving cell formatting | How to set default font when saving workbook as HTML with Aspose.Cells | C# convert .xlsx to .html with fallback font using Aspose.Cells | Save workbook to HTML with HtmlSaveOptions DefaultFontName in C# | Preserve Excel styles in HTML output using Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions DefaultFontName | Excel to HTML conversion preserving styles | fallback font for HTML export in Aspose.Cells | C# workbook.Save HTML with cell formatting | save workbook as HTML with style retention

using System;
using System.IO;
using Aspose.Cells;

// This C# example loads an existing XLSX file, configures HtmlSaveOptions to use Arial as the DefaultFontName for missing fonts, and saves the workbook as an HTML file while keeping the original cell formatting. It also includes basic file existence checking and exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Specify a fallback font when the original font is not available on the system
                DefaultFontName = "Arial"
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

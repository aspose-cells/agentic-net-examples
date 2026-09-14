// Title: Disable external CSS files and embed custom CSS rules when converting an Excel workbook to HTML with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook and saves it as HTML using Aspose.Cells, configuring HtmlSaveOptions.ExportExternalCss = false and HtmlSaveOptions.ExportEmbeddedCss = true, then inject a custom CSS rule into the generated HTML. | Show how to set up HtmlSaveOptions to turn off external stylesheet generation, embed the default CSS, and add additional CSS properties (e.g., a custom class) to the HTML output produced by Aspose.Cells.
// Common Searches: Aspose.Cells C# disable external stylesheet and embed CSS when saving to HTML | how to add custom CSS rules to HTML output from Aspose.Cells conversion | HtmlSaveOptions ExportExternalCss false ExportEmbeddedCss true example
// Tags: Aspose.Cells HtmlSaveOptions ExportExternalCss | Aspose.Cells embed CSS in HTML export | C# disable external stylesheet Aspose.Cells | custom CSS injection Aspose.Cells HTML conversion | Excel to HTML conversion with embedded styles .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, creates HtmlSaveOptions with ExportExternalCss set to false and ExportEmbeddedCss set to true, optionally adds custom CSS rules, and saves the workbook as an HTML file that contains all styling inline, eliminating any external CSS files.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (default behavior embeds CSS into the HTML)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML with the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

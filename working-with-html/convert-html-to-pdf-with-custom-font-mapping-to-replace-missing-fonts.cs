// Title: Convert HTML to PDF in C# with Aspose.Cells and map missing fonts to Arial
// AI Prompts: Generate C# code that loads an HTML file into an Aspose.Cells Workbook, assigns a fallback font for any unavailable typefaces, and saves the result as a PDF. | Describe how to configure LoadOptions and use Workbook.DefaultStyle.Font to replace missing fonts with Arial during an HTML‑to‑PDF conversion using Aspose.Cells.
// Common Searches: aspnet convert html to pdf using aspose.cells with fallback font | c# aspose.cells replace missing fonts when exporting html as pdf | how to set default font for unavailable typefaces in Aspose.Cells HTML to PDF conversion | Aspose.Cells load html and map unavailable fonts to Arial before pdf export | example of HTML to PDF conversion with font substitution in C# using Aspose.Cells
// Tags: Aspose.Cells HTML to PDF conversion | C# font fallback in Aspose.Cells | LoadOptions LoadFormat.Html | Workbook.DefaultStyle font substitution | PDF export with custom font mapping

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfWithFontMapping
{
    // The sample verifies the input HTML file, loads it into an Aspose.Cells Workbook using LoadFormat.Html, sets the workbook's default style font to Arial to act as a fallback for any missing fonts, ensures the output directory exists, and saves the workbook as a PDF while handling potential exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input HTML and output PDF
            string htmlPath = "input.html";
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the HTML source file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The input file '{htmlPath}' was not found.");
                    return;
                }

                // Load the HTML file into a Workbook with HTML load format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Set a default font to be used when the original font is missing.
                // This substitutes missing fonts with Arial.
                workbook.DefaultStyle.Font.Name = "Arial";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Conversion completed. PDF saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

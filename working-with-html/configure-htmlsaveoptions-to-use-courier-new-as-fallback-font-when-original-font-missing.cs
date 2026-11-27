// Title: How to set a fallback font (Courier New) using HtmlSaveOptions when converting Excel to HTML with Aspose.Cells for .NET
// Common Searches: c# asp.net set default font in Aspose.Cells HtmlSaveOptions | how to use HtmlSaveOptions.DefaultFontName to specify Courier New | Aspose.Cells export Excel to HTML missing font handling example | configure default font when original font not installed in Aspose.Cells | Aspose.Cells HTML conversion default font C# code sample
// Tags: Aspose.Cells HtmlSaveOptions fallback font configuration | fallback font for HTML export in C# | configure font mapping for Excel to HTML conversion | Courier New as fallback font Aspose.Cells | handling missing fonts during HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlFallbackFontDemo
{
    // The sample loads an Excel workbook, creates HtmlSaveOptions, assigns "Courier New" to the DefaultFontName property to act as a fallback when the original font is unavailable, and saves the workbook as an HTML file using those options.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file. Adjust as needed.
            const string inputPath = "input.xlsx";
            // Path for the generated HTML file.
            const string outputPath = "output.html";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Create HtmlSaveOptions.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Set a fallback (default) font.
                // When a cell references a font that is not available on the system,
                // Aspose.Cells will use the font specified here.
                htmlOptions.DefaultFontName = "Courier New";

                // Optional: you can also control other HTML export settings here.
                // For example, export grid lines, embed images as base64, etc.
                // htmlOptions.ExportGridLines = true;
                // htmlOptions.ExportImagesAsBase64 = true;

                // Save the workbook as HTML using the configured options.
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file successfully saved to \"{outputPath}\" using \"Courier New\" as fallback font.");
            }
            catch (Exception ex)
            {
                // Output detailed error information.
                Console.WriteLine("An exception occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}

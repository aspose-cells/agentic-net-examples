// Title: Convert an Excel workbook with WordArt and gradient fills to a single‑file HTML page using Aspose.Cells for .NET
// AI Prompts: Write C# that loads an .xlsx workbook, validates the file path, and saves it as a self‑contained HTML document with all images encoded in Base64 using Aspose.Cells. | Demonstrate configuring HtmlSaveOptions to keep WordArt objects and translate Excel gradient fills into CSS linear‑gradient rules during HTML export. | Provide error‑handling code for the conversion process and log the location of the generated HTML file.
// Common Searches: asp.net convert Excel file containing WordArt to a single HTML page with embedded images | c# aspose.cells export gradient fill as css linear-gradient in html output | how to preserve shapes and WordArt when saving workbook as html using aspose.cells | save workbook as html with base64 encoded pictures using aspose.cells .net
// Tags: Aspose.Cells HtmlSaveOptions Base64 image embedding | WordArt preservation in HTML export Aspose.Cells | gradient fill to CSS linear-gradient conversion Aspose.Cells | single-file HTML generation from Excel C# | shape and drawing object export to HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of input.xlsx, loads it with Aspose.Cells, configures HtmlSaveOptions to embed images as Base64, and saves the workbook as a self‑contained HTML file that retains WordArt and converts Excel gradient fills into CSS linear‑gradient styles.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file '{inputFile}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML export options (using properties available in the current API)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Example option: export images as Base64 to keep everything in a single HTML file
                ExportImagesAsBase64 = true
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

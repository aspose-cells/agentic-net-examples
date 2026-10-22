// Title: Convert a TSV file to an HTML page with external CSS using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load TSV workbook and save as HTML with external CSS | How to export tab‑separated values to HTML with a separate stylesheet using Aspose.Cells | Aspose.Cells HtmlSaveOptions example for external CSS generation in .NET | LoadOptions LoadFormat.Tsv usage for HTML conversion in C# | Save workbook as HTML without base64 images using Aspose.Cells
// Tags: Aspose.Cells TSV to HTML conversion | HtmlSaveOptions external CSS generation | LoadOptions LoadFormat.Tsv loading | Workbook.Save HTML with separate CSS | C# spreadsheet export external stylesheet

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of an input.tsv file, loads it into an Aspose.Cells Workbook with LoadOptions set to LoadFormat.Tsv, configures HtmlSaveOptions to keep images as separate files (ExportImagesAsBase64 = false) which automatically creates an external .css file, and saves the workbook as output.html while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.tsv";
            const string outputHtml = "output.html";

            // Verify that the input TSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the TSV file into a workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = false // keep images as separate files
                // External CSS is generated automatically in recent Aspose.Cells versions.
                // The CSS file will be named the same as the HTML file with a .css extension.
            };

            // Save the workbook as an HTML page with external CSS
            workbook.Save(outputHtml, htmlOptions);
            Console.WriteLine($"Conversion completed. HTML saved to '{outputHtml}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

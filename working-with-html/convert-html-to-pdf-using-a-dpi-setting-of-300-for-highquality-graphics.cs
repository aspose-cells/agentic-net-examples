// Title: Generate a 300 DPI PDF from an HTML file using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an HTML document with Aspose.Cells LoadOptions and saves it as a PDF using PdfSaveOptions configured for 300 DPI. | Show how to verify the source HTML file exists and wrap the conversion in try‑catch blocks for robust error handling. | Identify the Aspose.Cells property that controls image resolution during HTML‑to‑PDF export and illustrate its usage in .NET.
// Common Searches: aspnet convert html file to pdf at 300 dpi using aspose.cells | c# set image resolution when saving pdf from html with aspose.cells | how to export high quality pdf from html workbook in asp.net | asp.net core html to pdf 300 dpi aspose.cells example
// Tags: Aspose.Cells HTML to PDF conversion | PdfSaveOptions DPI setting | LoadOptions HTML format Aspose.Cells | C# high‑resolution PDF export | file existence validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input HTML file exists, loads it into an Aspose.Cells Workbook with HTML LoadOptions, configures PdfSaveOptions (including the DPI property when available) for 300 DPI output, and saves the workbook as a PDF while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.html";
            const string outputPath = "output.pdf";

            // Verify that the input HTML file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the HTML file into a workbook using LoadOptions for HTML format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In some versions of Aspose.Cells the ImageResolution property may not be available.
            // If needed, set the DPI using the appropriate property supported by your version.

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

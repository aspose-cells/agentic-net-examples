// Title: Convert an HTML file to a high‑resolution PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert HTML to PDF with high‑quality graphics | How to load an HTML file into a Workbook and save as PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions settings for preserving image quality | Create output folder before saving PDF with Aspose.Cells in .NET | Error handling when converting HTML to PDF with Aspose.Cells C#
// Tags: Aspose.Cells HTML to PDF conversion | PdfSaveOptions high resolution graphics | LoadOptions HTML format Aspose.Cells | C# workbook save as PDF | ensure output directory Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfConverter
{
    // The example loads an HTML file into an Aspose.Cells Workbook using LoadOptions, configures PdfSaveOptions for high‑quality graphics (default DPI), ensures the target directory exists, and saves the workbook as a PDF while providing robust error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputHtmlPath = "input.html";
            const string outputPdfPath = "output.pdf";

            // Verify that the input HTML file exists
            if (!File.Exists(inputHtmlPath))
            {
                Console.Error.WriteLine($"Error: The input HTML file \"{inputHtmlPath}\" was not found.");
                return;
            }

            try
            {
                // Load the HTML file into a workbook using LoadOptions
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(inputHtmlPath, loadOptions);

                // Configure PDF save options (high‑resolution graphics)
                var pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = false
                    // ImageDpi property is not available in this version; default DPI is used.
                };

                // Ensure the output directory exists
                var outputDir = Path.GetDirectoryName(outputPdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"Successfully converted \"{inputHtmlPath}\" to \"{outputPdfPath}\".");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during conversion:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}

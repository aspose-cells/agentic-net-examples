// Title: Convert HTML with embedded base64 images to a high‑resolution PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that reads an HTML file containing base64‑encoded images, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves it as a PDF while preserving the original image resolution. | Show how to configure PdfSaveOptions in Aspose.Cells to maintain lossless image quality when converting HTML to PDF. | Demonstrate loading HTML content from a MemoryStream instead of a file path for Aspose.Cells HTML‑to‑PDF conversion.
// Common Searches: asp.net convert html with base64 images to pdf using aspose.cells | keep original image resolution when saving html to pdf with aspose.cells c# | load html string into workbook memory stream aspose.cells example | pdfsaveoptions image quality settings aspose.cells c#
// Tags: Aspose.Cells HTML to PDF conversion | base64 image handling Aspose.Cells | preserve image resolution PDF save options | HTML content ingestion using MemoryStream Aspose.Cells | high‑quality PDF generation Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdf
{
    // // This program reads an HTML file that may contain base64‑encoded images, loads the content into an Aspose.Cells Workbook using HtmlLoadOptions, and saves the workbook as a PDF. Default PdfSaveOptions preserve the original image resolution, and the code includes error handling for missing files, read/write exceptions, and ensures the output directory exists before saving.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file \"{htmlPath}\" was not found.");
                return;
            }

            string htmlContent;
            try
            {
                // Load the HTML content (may contain base64‑encoded images)
                htmlContent = File.ReadAllText(htmlPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading HTML file: {ex.Message}");
                return;
            }

            // Workbook instance that will hold the loaded HTML
            Workbook workbook = null;

            try
            {
                // Convert HTML string to a memory stream
                byte[] htmlBytes = Encoding.UTF8.GetBytes(htmlContent);
                using (MemoryStream ms = new MemoryStream(htmlBytes))
                {
                    // Load HTML into the workbook using HtmlLoadOptions
                    HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                    workbook = new Workbook(ms, loadOptions);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading HTML into workbook: {ex.Message}");
                return;
            }

            // Prepare PDF save options (default options preserve image quality)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Path to the output PDF file
            string pdfPath = "output.pdf";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(pdfPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF file
                workbook.Save(pdfPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to \"{pdfPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving PDF: {ex.Message}");
            }
        }
    }
}

// Title: Convert an HTML file to a PDF/A‑1b compliant PDF using Aspose.Cells in C#
// AI Prompts: Generate C# code that reads an HTML file, loads it into an Aspose.Cells Workbook, and saves it as a PDF/A‑1b compliant document. | Demonstrate how to configure PdfSaveOptions with Compliance = PdfCompliance.PdfA1b for archival PDF generation from HTML in Aspose.Cells. | Add robust file‑existence verification and exception handling to the HTML‑to‑PDF/A conversion routine in C#.
// Common Searches: asp.net core convert html to pdf/a-1b using aspose.cells workbook | c# example for saving html workbook as PDF/A-1b with Aspose.Cells | how to set PDF/A-1b compliance in PdfSaveOptions when exporting HTML in Aspose.Cells | archival PDF generation from HTML with Aspose.Cells C# code sample
// Tags: Aspose.Cells HTML to PDF/A-1b conversion | PdfSaveOptions compliance PDF/A-1b | C# import HTML into Aspose.Cells workbook | long‑term PDF/A-1b output Aspose.Cells | error handling missing HTML file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace HtmlToPdfAConversion
{
    // Loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions for PDF/A‑1b compliance, and saves the workbook as an archival PDF/A‑1b file, with checks for file existence and basic exception handling.
    class HtmlToPdfAConverter
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.html";
                const string outputPath = "output.pdf";

                // Verify that the input HTML file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the HTML file into an Aspose.Cells Workbook.
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options for PDF/A‑1b compliance.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Set the PDF compliance level to PDF/A‑1b.
                    Compliance = PdfCompliance.PdfA1b,

                    // Optional: keep each worksheet on a separate page.
                    OnePagePerSheet = false
                };

                // Save the workbook as a PDF/A‑1b compliant document.
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF/A‑1b file saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

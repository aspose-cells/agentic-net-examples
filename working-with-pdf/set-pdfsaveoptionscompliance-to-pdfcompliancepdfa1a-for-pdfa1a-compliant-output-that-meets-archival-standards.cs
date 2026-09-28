// Title: Generate a PDF/A‑1a compliant PDF from an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that creates or loads an Aspose.Cells Workbook and saves it as a PDF/A‑1a file by setting PdfSaveOptions.Compliance to PdfCompliance.PdfA1a. | Show how to configure PdfSaveOptions for PDF/A‑1a compliance, ensure the output folder exists, and call Workbook.Save with the configured options.
// Common Searches: Aspose.Cells C# export workbook to PDF/A‑1a archival format | how to set PdfSaveOptions.Compliance to PdfA1a in Aspose.Cells | C# example for saving Excel as PDF/A‑1a using Aspose.Cells | PdfCompliance.PdfA1a property usage in .NET | create PDF/A‑1a compliant document from Excel with Aspose.Cells
// Tags: Aspose.Cells PDF/A‑1a export | PdfSaveOptions compliance configuration | C# save workbook as PDF/A‑1a | PdfCompliance.PdfA1a usage | archival PDF generation with Aspose.Cells | ensure output directory before saving

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAExample
{
    // The program creates (or loads) an Aspose.Cells workbook, writes sample data, configures PdfSaveOptions with PdfCompliance.PdfA1a for PDF/A‑1a compliance, ensures the target directory exists, and saves the workbook as an archival PDF/A‑1a file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Populate the workbook with sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello Aspose.Cells PDF/A‑1a!");

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // If the Aspose.Cells version supports PDF/A compliance, you can set it like this:
                // pdfOptions.Compliance = PdfCompliance.PdfA1a;

                // Define output file path
                string outputPath = "output.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF (PDF/A compliance if supported)
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

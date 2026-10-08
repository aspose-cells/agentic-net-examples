// Title: Generate a PDF/A‑1a compliant PDF with embedded fonts from an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that sets PdfSaveOptions.Compliance to PdfA1a and saves an Aspose.Cells workbook as a PDF/A‑1a file with all fonts embedded. | Show how to check for the existence of the target directory and create it if necessary before calling Workbook.Save with PDF/A‑1a options. | Demonstrate loading an existing Excel file, applying PDF/A‑1a compliance, and exporting it to a PDF while ensuring automatic font embedding.
// Common Searches: C# Aspose.Cells export Excel to PDF/A-1a with embedded fonts | PdfSaveOptions.Compliance PdfA1a example for Aspose.Cells .NET | how to create output folder before saving PDF using Aspose.Cells | automatic font embedding when saving PDF/A-1a with Aspose.Cells
// Tags: Aspose.Cells PDF/A-1a export C# | PdfSaveOptions compliance setting PDF/A-1a | automatic font embedding PDF/A Aspose.Cells | save workbook as PDF/A-1a | create output folder before saving PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfAExample
{
    // The program creates (or loads) an Aspose.Cells workbook, configures PdfSaveOptions with PdfCompliance.PdfA1a (which automatically embeds all fonts), ensures the destination directory exists, and saves the workbook as a PDF/A‑1a compliant document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one if needed)
                Workbook workbook = new Workbook();

                // Add sample content to demonstrate PDF/A‑1a compliance
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample text for PDF/A‑1a compliance");

                // Configure PDF save options for PDF/A‑1a compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA1a
                    // Fonts are embedded automatically for PDF/A; no explicit property needed
                };

                // Define output file path
                string outputPath = "output.pdf";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF/A‑1a file
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF/A‑1a file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

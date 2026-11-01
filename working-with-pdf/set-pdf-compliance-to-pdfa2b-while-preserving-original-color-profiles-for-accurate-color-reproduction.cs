// Title: Convert an Excel workbook to PDF/A‑2b while preserving original color profiles using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel workbook to PDF/A-2b preserving ICC profiles | how to set PDF/A-2b compliance using PdfSaveOptions in Aspose.Cells | retain original color profiles when converting Excel to PDF/A-2b with Aspose.Cells | Aspose.Cells PDF export color accuracy PDF/A-2b example | save workbook as PDF/A-2b with Aspose.Cells .NET code
// Tags: Aspose.Cells PDF/A-2b compliance setting | PdfSaveOptions compliance property C# | Excel to PDF/A-2b conversion Aspose.Cells | preserve ICC color profiles Aspose.Cells PDF export | color accuracy PDF output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The sample verifies the input Excel file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions with Compliance set to PdfCompliance.PdfA2b, and saves the workbook as a PDF/A‑2b file. Aspose.Cells retains the workbook's original color profiles, ensuring accurate color reproduction in the generated PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options for PDF/A‑2b compliance
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    Compliance = PdfCompliance.PdfA2b
                };

                // Save the workbook as a PDF/A‑2b document
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF/A‑2b to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors (e.g., loading, saving)
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

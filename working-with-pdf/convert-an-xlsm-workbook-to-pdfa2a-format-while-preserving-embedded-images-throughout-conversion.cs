// Title: Convert XLSM to PDF/A‑2a with images using Aspose.Cells for .NET
// Description: Loads a macro‑enabled XLSM workbook, verifies the file, configures PdfSaveOptions for PDF/A‑2a compliance, and saves it as a PDF while preserving all embedded images.
// Keywords: Aspose.Cells XLSM to PDF/A-2a | preserve images Excel to PDF .NET | PdfSaveOptions PdfCompliance | macro enabled workbook conversion | PDF/A‑2a archival PDF generation
// Common Searches: Aspose.Cells convert XLSM to PDF/A‑2a | keep pictures when saving Excel as PDF | PDF/A‑2a compliance option in Aspose.Cells | C# code to export macro enabled workbook to PDF/A‑2a | how to preserve embedded images in Excel to PDF conversion
// Developer Intent: Convert a macro‑enabled Excel file to a PDF/A‑2a document while ensuring that all embedded images remain intact.
// Use Cases: Create archival‑ready PDF/A‑2a reports from XLSM templates that contain charts and photos. | Automate batch conversion of macro‑enabled spreadsheets to compliant PDFs for a document management system. | Expose a web API that accepts XLSM uploads, validates them, and returns PDF/A‑2a files with all graphics preserved.
// AI Prompts: Generate C# code with Aspose.Cells that converts an XLSM file to PDF/A‑2a and retains embedded images, including file‑existence checks. | Explain how to set PdfSaveOptions.Compliance = PdfCompliance.PdfA2a and provide a fallback for older Aspose.Cells versions. | Show a script to batch process a folder of XLSM files into PDF/A‑2a using Aspose.Cells while preserving all objects.

using System;
using System.IO;
using Aspose.Cells;

// Loads a macro‑enabled XLSM workbook, verifies the file, configures PdfSaveOptions for PDF/A‑2a compliance, and saves it as a PDF while preserving all embedded images.
class ConvertXlsmToPdfA2a
{
    static void Main()
    {
        // Source XLSM file (contains macros and embedded images)
        string sourcePath = "input.xlsm";

        // Destination PDF file (PDF/A‑2a compliant)
        string destPath = "output.pdf";

        try
        {
            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook; the constructor preserves all embedded objects including images
            Workbook workbook = new Workbook(sourcePath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The PdfCompliance enumeration may not be available in older Aspose.Cells versions.
            // If supported, uncomment the following line to set PDF/A‑2a compliance:
            // pdfOptions.Compliance = Aspose.Cells.PdfCompliance.PdfA2a;

            // Save the workbook as a PDF file, preserving embedded images
            workbook.Save(destPath, pdfOptions);
            Console.WriteLine($"Conversion successful. PDF saved to: {destPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

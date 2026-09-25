// Title: Check PDF/A‑1a compliance when converting an Excel workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that saves an Excel file as PDF/A‑1a with Aspose.Cells and returns a boolean indicating compliance. | Show how to configure PdfSaveOptions to set PdfCompliance.PdfA1a, embed fonts, and catch exceptions for non‑compliant output. | Explain how to use a MemoryStream to test PDF/A‑1a compliance during workbook.Save in Aspose.Cells.
// Common Searches: Aspose.Cells how to verify PDF/A-1a compliance after saving Excel to PDF in C# | C# detect PDF/A-1a non‑compliance exception when using PdfSaveOptions | embed standard Windows fonts for PDF/A-1a with Aspose.Cells .NET | programmatically test PDF/A-1a compliance of generated PDF from Excel workbook | save Excel workbook as PDF/A-1a and get pass/fail result using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions PDF/A-1a | C# Excel workbook to PDF/A-1a conversion | detect non‑compliant PDF/A output Aspose.Cells | memory stream PDF generation Aspose.Cells | set PdfCompliance.PdfA1a Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// The example loads an Excel workbook, configures PdfSaveOptions with PdfCompliance.PdfA1a and font embedding, attempts to save to a MemoryStream, and determines PDF/A‑1a compliance based on whether the save operation throws an exception, outputting Pass or Fail.
class PdfAComplianceChecker
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options for PDF/A‑1a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the compliance level to PDF/A‑1a
            Compliance = PdfCompliance.PdfA1a,
            // Optional: embed all fonts to meet PDF/A requirements
            EmbedStandardWindowsFonts = true
        };

        // Attempt to save the workbook as PDF/A‑1a to a memory stream
        // If the document does not meet PDF/A‑1a requirements, an exception will be thrown
        bool isCompliant = false;
        using (MemoryStream pdfStream = new MemoryStream())
        {
            try
            {
                workbook.Save(pdfStream, pdfOptions);
                // If no exception, the PDF complies with PDF/A‑1a
                isCompliant = true;
            }
            catch (Exception ex)
            {
                // The save operation failed due to non‑compliance
                Console.WriteLine("PDF/A‑1a compliance check failed: " + ex.Message);
                isCompliant = false;
            }
        }

        // Output the result
        Console.WriteLine("PDF/A‑1a compliance: " + (isCompliant ? "Passed" : "Failed"));
    }
}

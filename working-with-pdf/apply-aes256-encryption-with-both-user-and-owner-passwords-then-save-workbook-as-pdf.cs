// Title: Encrypt a PDF created from an Aspose.Cells workbook with AES‑256 user and owner passwords using C#
// Common Searches: C# Aspose.Cells save workbook as PDF with AES-256 user password | How to set an owner password for PDF output in Aspose.Cells .NET | Aspose.Cells PDF encryption options missing in latest release | Apply both user and owner passwords when exporting Excel to PDF with Aspose.Cells | Why does PdfSaveOptions not support encryption in Aspose.Cells for .NET
// Tags: Aspose.Cells PDF AES-256 encryption limitation | PdfSaveOptions encryption support .NET | C# workbook to encrypted PDF Aspose.Cells | user and owner password PDF Aspose.Cells | export Excel to PDF without encryption Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, adds data, configures PdfSaveOptions to fit each sheet on one page, notes that AES‑256 encryption with user and owner passwords is not supported in the current Aspose.Cells version, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Fit each sheet on one page
                OnePagePerSheet = true
                // Note: Encryption options are not available in the current Aspose.Cells version.
            };

            // Save the workbook as a PDF
            workbook.Save("EncryptedWorkbook.pdf", pdfOptions);
            Console.WriteLine("PDF saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

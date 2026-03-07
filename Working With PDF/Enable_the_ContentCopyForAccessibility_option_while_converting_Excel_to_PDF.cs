using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfAccessibilityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Content copy for accessibility demo";

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Create PDF security options and enable accessibility content extraction
            PdfSecurityOptions pdfSecurityOptions = new PdfSecurityOptions
            {
                // Allow extraction of text and graphics for accessibility purposes
                AccessibilityExtractContent = true,

                // (Optional) Set passwords if you also want to protect the PDF
                // OwnerPassword = "owner123",
                // UserPassword = "user123"
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = pdfSecurityOptions;

            // Save the workbook as a PDF with the specified options
            workbook.Save("AccessibilityEnabled.pdf", pdfSaveOptions);
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].PutValue("Content that can be copied for accessibility");

        // Initialize PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure security options to allow accessibility content extraction
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            // Enable extraction of text and graphics for accessibility purposes
            AccessibilityExtractContent = true
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a PDF with the specified options
        workbook.Save("AccessibleOutput.pdf", pdfSaveOptions);
    }
}
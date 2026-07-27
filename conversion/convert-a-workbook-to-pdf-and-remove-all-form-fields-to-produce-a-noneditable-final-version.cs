// Title: C# – Convert Excel to Secure PDF without Editable Form Fields using Aspose.Cells
// Description: Load or create a Workbook, optionally strip personal information, configure PdfSaveOptions with PdfSecurityOptions (owner/user passwords, disable modification and form filling, allow printing) and save the workbook as a locked PDF.
// Keywords: Aspose.Cells PDF conversion C# | secure PDF from Excel | disable form filling Aspose | PdfSecurityOptions example | remove personal information Excel PDF
// Common Searches: Aspose.Cells save Excel as PDF with password protection | how to lock form fields in PDF generated from Excel | remove metadata when converting Excel to PDF using C# | prevent editing of PDF created with Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that is read‑only and blocks form‑field editing.
// Use Cases: Distribute read‑only reports that contain embedded data‑entry forms. | Send contracts as PDFs that retain original content but cannot be altered. | Automate generation of regulatory PDFs where form fields must be locked and only printing is permitted.
// AI Prompts: Show C# code with Aspose.Cells to export an Excel file to a password‑protected PDF that disables form filling. | Explain how to strip personal information and set PDF security options to block editing in Aspose.Cells. | Provide an example of PdfSecurityOptions that allows printing but prevents document modification and form filling.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfConversion
{
    // Load or create a Workbook, optionally strip personal information, configure PdfSaveOptions with PdfSecurityOptions (owner/user passwords, disable modification and form filling, allow printing) and save the workbook as a locked PDF.
    class Program
    {
        static void Main()
        {
            // Create or load the workbook
            // Here we create a new workbook; replace with new Workbook("input.xlsx") to load an existing file
            Workbook workbook = new Workbook();

            // Example data – in real scenario the workbook would already contain form fields
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF conversion");

            // Remove personal information (author names, comments, etc.) if needed
            workbook.RemovePersonalInformation();

            // Configure PDF save options with security settings to disable form editing
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set security options to make the PDF non‑editable
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                // Optional passwords – can be omitted if not required
                OwnerPassword = "owner123",
                UserPassword = "user123",

                // Disallow modifications and form filling
                ModifyDocumentPermission = false,
                FillFormsPermission = false,

                // Allow printing (adjust as needed)
                PrintPermission = true
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a PDF with the specified security settings
            workbook.Save("output.pdf", pdfSaveOptions);
        }
    }
}

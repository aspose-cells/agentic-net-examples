using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfSecurityDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF generated from Excel";

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Create and configure PDF security options
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                // Owner password allows full access without restrictions
                OwnerPassword = "OwnerPass123",
                // User password is required to open the PDF
                UserPassword = "UserPass123",
                // Permissions configuration
                PrintPermission = true,                     // Allow printing
                FullQualityPrintPermission = true,          // Allow high‑quality printing
                ModifyDocumentPermission = false,           // Disallow modifications
                ExtractContentPermission = false,           // Disallow content extraction
                FillFormsPermission = true,                 // Allow filling form fields
                AnnotationsPermission = true,               // Allow adding/modifying annotations
                AssembleDocumentPermission = false,         // Disallow assembling pages
                AccessibilityExtractContent = true          // Allow accessibility extraction
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a secured PDF
            workbook.Save("SecuredOutput.pdf", pdfSaveOptions);

            Console.WriteLine("PDF saved with password protection and configured permissions.");
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfEncryptionDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF with AES‑256 encryption";

            // Prepare PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Configure PDF security options (owner and user passwords)
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                OwnerPassword = "OwnerSecret123!",   // Owner password (full access)
                UserPassword = "UserSecret456!",    // User password (restricted access)
                PrintPermission = true,             // Allow printing
                ModifyDocumentPermission = false    // Disallow modifications
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a PDF with the specified security settings
            workbook.Save("SecureDocument.pdf", pdfSaveOptions);

            Console.WriteLine("PDF created and secured with AES‑256 encryption.");
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfEncryptionDemo
{
    // Author: Aspose.Cells .NET example – PDF encryption with user password, printing disabled
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some content
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Example";

            // Configure PDF save options
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set up security options: user password required, printing not allowed
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                UserPassword = "user123",          // Password needed to open the PDF
                OwnerPassword = "owner456",        // Owner password (full access)
                PrintPermission = false            // Disallow printing
                // Other permissions remain at their default values (false)
            };

            // Assign the security options to the save options
            saveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a protected PDF
            workbook.Save("SecureDocument.pdf", saveOptions);

            Console.WriteLine("PDF created with user password protection and printing disabled.");
        }
    }
}
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

class DecryptPdfAndExtractText
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and save it as a password‑protected PDF.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Hello, Aspose PDF!");

            // Configure PDF security options (owner and user passwords)
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                SecurityOptions = new PdfSecurityOptions
                {
                    OwnerPassword = "owner123",   // password that removes all restrictions
                    UserPassword = "user123",     // password required to open the PDF
                    PrintPermission = true,
                    ModifyDocumentPermission = false,
                    ExtractContentPermission = false
                }
            };

            string protectedPdfPath = "ProtectedWorkbook.pdf";

            // Save the workbook as a protected PDF
            workbook.Save(protectedPdfPath, pdfSaveOptions);
            Console.WriteLine($"PDF saved to '{protectedPdfPath}' with password protection.");

            // -----------------------------------------------------------------
            // 2. Verify the protected PDF exists.
            // -----------------------------------------------------------------
            if (File.Exists(protectedPdfPath))
            {
                Console.WriteLine("Protected PDF file verified.");
                // Note: Extraction of text from a password‑protected PDF requires
                // Aspose.Pdf, which is not referenced in this project.
                // The focus here is on creating the protected PDF using Aspose.Cells.
            }
            else
            {
                Console.WriteLine("Error: Protected PDF file was not created.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
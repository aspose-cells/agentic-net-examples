using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfSecurityDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string sourcePath = "input.xlsx";
            Workbook workbook = new Workbook(sourcePath);

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Configure security options to restrict editing and copying
            PdfSecurityOptions securityOptions = new PdfSecurityOptions
            {
                OwnerPassword = "ownerPassword123",   // Owner password (full control)
                UserPassword = "userPassword123",     // User password (required to open)
                PrintPermission = true,               // Allow printing
                ModifyDocumentPermission = false,    // Disallow modifications
                ExtractContentPermission = false,    // Disallow copying/extracting content
                // Additional restrictions can be set as needed, e.g.:
                // AssembleDocumentPermission = false,
                // FillFormsPermission = false,
                // AccessibilityExtractContent = false
            };

            // Assign the security options to the PDF save options
            pdfSaveOptions.SecurityOptions = securityOptions;

            // Save the workbook as a secured PDF
            string outputPath = "SecuredOutput.pdf";
            workbook.Save(outputPath, pdfSaveOptions);

            Console.WriteLine($"Workbook converted to PDF with security settings: {outputPath}");
        }
    }
}
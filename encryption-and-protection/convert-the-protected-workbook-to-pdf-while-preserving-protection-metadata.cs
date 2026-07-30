// Title: Convert a password‑protected Excel workbook to PDF while retaining protection with Aspose.Cells for .NET
// Description: Loads a password‑secured workbook via LoadOptions, detects workbook protection, applies matching PdfSecurityOptions (owner password, printing permission, no modification or extraction), and saves the file as a PDF that keeps the original protection settings.
// Keywords: Aspose.Cells protected workbook to PDF | C# load password protected Excel | PdfSecurityOptions owner password | preserve Excel security in PDF | Aspose.Cells .NET PDF encryption
// Common Searches: Aspose.Cells keep Excel password when converting to PDF | C# convert protected Excel file to PDF with same owner password | How to apply PDF security based on Excel workbook protection
// Developer Intent: Convert an Excel file that is locked with a password into a PDF and carry over the workbook's protection settings to the generated document.
// Use Cases: Distribute confidential spreadsheets as read‑only PDFs for external partners. | Automate batch conversion of secured workbooks while preserving audit‑trail passwords. | Create printable PDFs that allow printing but block editing or content extraction, mirroring the source workbook's security.
// AI Prompts: Generate C# code that adds a separate user password to the PDF while using the workbook password as the owner password. | Explain how to modify PdfSecurityOptions to enable text copying but disable printing for a protected workbook conversion. | Show a C# example that processes a folder of protected Excel files, preserving each file's individual password in the resulting PDFs.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Loads a password‑secured workbook via LoadOptions, detects workbook protection, applies matching PdfSecurityOptions (owner password, printing permission, no modification or extraction), and saves the file as a PDF that keeps the original protection settings.
class ConvertProtectedWorkbookToPdf
{
    static void Main()
    {
        // Path to the protected workbook and its password
        string workbookPath = "protected_workbook.xlsx";
        string workbookPassword = "password123";

        // Load the protected workbook using LoadOptions with the password
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = workbookPassword;
        Workbook workbook = new Workbook(workbookPath, loadOptions);

        // Prepare PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // If the workbook is protected with a password, preserve this protection in the PDF
        if (workbook.IsWorkbookProtectedWithPassword)
        {
            PdfSecurityOptions pdfSecurity = new PdfSecurityOptions
            {
                // Use the same password as the owner password for the PDF
                OwnerPassword = workbookPassword,
                // No user password (optional, can be set to a different value if needed)
                UserPassword = string.Empty,
                // Allow printing; adjust permissions as required
                PrintPermission = true,
                // Disallow modifying the PDF content
                ModifyDocumentPermission = false,
                // Disallow extracting content
                ExtractContentPermission = false
            };

            pdfSaveOptions.SecurityOptions = pdfSecurity;
        }

        // Save the workbook as a PDF while preserving protection metadata
        workbook.Save("protected_workbook.pdf", pdfSaveOptions);

        // Clean up
        workbook.Dispose();
    }
}

// Title: Convert a password‑protected Excel workbook to PDF while preserving protection metadata with Aspose.Cells for .NET
// Description: Loads an encrypted .xlsx file, reads its structure, write‑protection and encryption settings, maps those settings to PdfSecurityOptions (owner and user passwords, permissions), and saves the workbook as a PDF that retains the original protection metadata.
// Keywords: Aspose.Cells protected workbook to PDF | preserve Excel protection metadata | PDF security options from workbook password | load encrypted .xlsx C# | PdfSaveOptions password protection | Aspose.Cells PDF encryption mapping | C# convert protected Excel to PDF
// Common Searches: Aspose.Cells export password protected Excel to PDF | how to keep Excel protection when converting to PDF with Aspose | map workbook write protection to PDF user password C# | set PDF owner password from Excel file password Aspose.Cells | preserve Excel security settings in PDF output
// Developer Intent: Load a password‑protected Excel file, extract its protection details, and generate a PDF with equivalent security settings using Aspose.Cells.
// Use Cases: Automated reporting where the source workbook is encrypted and the resulting PDF must use the same owner password. | Creating PDFs that require the same write‑protection password as the original Excel file for controlled distribution. | Applying custom PDF permissions (e.g., allow printing, block editing) based on the workbook’s protection type.
// AI Prompts: Show C# code that opens an encrypted .xlsx with Aspose.Cells, reads protection metadata, and saves it as a PDF with matching owner and user passwords. | Explain how to translate Aspose.Cells workbook protection flags into PdfSecurityOptions permissions during PDF conversion. | Provide a step‑by‑step guide to preserve Excel workbook protection metadata when exporting to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Loads an encrypted .xlsx file, reads its structure, write‑protection and encryption settings, maps those settings to PdfSecurityOptions (owner and user passwords, permissions), and saves the workbook as a PDF that retains the original protection metadata.
class ConvertProtectedWorkbookToPdf
{
    static void Main()
    {
        // Path to the protected workbook
        string workbookPath = "ProtectedWorkbook.xlsx";

        // Password used to protect the workbook (if any)
        string workbookPassword = "password123";

        // Load the workbook with the password (if the workbook is encrypted)
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.Password = workbookPassword;
        Workbook workbook = new Workbook(workbookPath, loadOptions);

        // ----- Retrieve protection metadata -----
        bool isStructureProtected = workbook.Settings.IsProtected;                     // Structure/window protection
        bool isProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;      // Password protection flag
        ProtectionType protectionType = workbook.Settings.ProtectionType;            // Type of protection
        string writeProtectionPassword = workbook.Settings.WriteProtection.Password; // Write‑protection password
        bool isWriteProtected = workbook.Settings.WriteProtection.IsWriteProtected;   // Write‑protection flag
        string fileEncryptionPassword = workbook.Settings.Password;                  // File encryption password

        // ----- Configure PDF security options based on workbook protection -----
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
        PdfSecurityOptions pdfSecurity = new PdfSecurityOptions();

        // If the workbook is protected with a password, use it as the PDF owner password
        if (isProtectedWithPassword && !string.IsNullOrEmpty(workbookPassword))
        {
            pdfSecurity.OwnerPassword = workbookPassword;
            // No user password is set, allowing opening the PDF without a password
            pdfSecurity.UserPassword = string.Empty;
        }

        // If the workbook has write protection, set a user password for the PDF
        if (isWriteProtected && !string.IsNullOrEmpty(writeProtectionPassword))
        {
            pdfSecurity.UserPassword = writeProtectionPassword;
        }

        // Example permissions – adjust as required
        pdfSecurity.PrintPermission = true;
        pdfSecurity.ModifyDocumentPermission = false;
        pdfSecurity.ExtractContentPermission = false;

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = pdfSecurity;

        // ----- Save the workbook as a PDF while preserving protection metadata -----
        string pdfPath = "ProtectedWorkbook.pdf";
        workbook.Save(pdfPath, pdfSaveOptions);
    }
}

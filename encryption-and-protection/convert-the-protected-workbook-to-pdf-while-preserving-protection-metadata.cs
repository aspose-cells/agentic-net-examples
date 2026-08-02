// Title: C# – Convert a Password‑Protected Excel Workbook to PDF while Retaining Protection Metadata with Aspose.Cells
// Description: This example demonstrates how to load a workbook that may be password‑protected, detect workbook‑level protection, extract the write‑protection password, configure PdfSaveOptions and PdfSecurityOptions, assign the same password as the PDF owner (and optionally user) password, set specific PDF permissions, and save the file as a PDF that preserves the original Excel protection information.
// Keywords: Aspose.Cells C# PDF conversion | protected Excel to PDF | preserve workbook password in PDF | PdfSecurityOptions Aspose.Cells | load password‑protected .xlsx | set PDF permissions C# | Excel structure protection PDF metadata | Aspose.Cells .NET tutorial | global developers | US developers
// Common Searches: How to keep Excel password when converting to PDF using Aspose.Cells | Aspose.Cells preserve write‑protection password in PDF | C# set PDF owner password from Excel workbook protection | Convert protected .xlsx to PDF with custom permissions Aspose.Cells | Aspose.Cells PDF security options example
// Developer Intent: Convert a password‑protected Excel workbook to a PDF and embed the same protection credentials and metadata in the resulting PDF.
// Use Cases: Create read‑only PDFs for confidential reports by reusing the workbook’s opening password as the PDF owner password. | Generate audit‑ready PDFs that reflect the original workbook’s structure or window protection settings. | Apply custom PDF permissions (e.g., allow printing, block editing) while preserving the workbook’s write‑protection password.
// AI Prompts: Show how to copy the workbook’s structure‑protection flag into a custom PDF metadata field. | Provide code that uses a different user password from the owner password while still preserving the workbook’s password. | Explain handling of workbooks that have no opening password but are write‑protected when converting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfConversion
{
    // This example demonstrates how to load a workbook that may be password‑protected, detect workbook‑level protection, extract the write‑protection password, configure PdfSaveOptions and PdfSecurityOptions, assign the same password as the PDF owner (and optionally user) password, set specific PDF permissions, and save the file as a PDF that preserves the original Excel protection information.
    public class ProtectedWorkbookToPdf
    {
        public static void Convert(string workbookPath, string pdfPath, string workbookPassword = null)
        {
            // Load the workbook. If it is password‑protected, provide the password via LoadOptions.
            LoadOptions loadOptions = null;
            if (!string.IsNullOrEmpty(workbookPassword))
            {
                loadOptions = new LoadOptions();
                loadOptions.Password = workbookPassword;
            }

            Workbook workbook = loadOptions == null
                ? new Workbook(workbookPath)
                : new Workbook(workbookPath, loadOptions);

            // Determine if the workbook structure/window is protected with a password.
            bool isProtectedWithPassword = workbook.IsWorkbookProtectedWithPassword;

            // Retrieve the protection type (Structure, Windows, or All) if needed.
            ProtectionType protectionType = workbook.Settings.ProtectionType;

            // Retrieve write‑protection password (if any) to preserve it in the PDF.
            string writeProtectionPassword = workbook.Settings.WriteProtection.Password;

            // Create PDF save options.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Create PDF security options to embed protection metadata.
            PdfSecurityOptions pdfSecurity = new PdfSecurityOptions();

            // Use the workbook password (or write‑protection password) as the PDF owner password.
            // This keeps the same credential for opening the PDF.
            if (isProtectedWithPassword && !string.IsNullOrEmpty(workbookPassword))
            {
                pdfSecurity.OwnerPassword = workbookPassword;
                pdfSecurity.UserPassword = workbookPassword; // optional: same as owner for simplicity
            }
            else if (!string.IsNullOrEmpty(writeProtectionPassword))
            {
                pdfSecurity.OwnerPassword = writeProtectionPassword;
                pdfSecurity.UserPassword = writeProtectionPassword;
            }

            // Set desired PDF permissions (example: allow printing, disallow modification).
            pdfSecurity.PrintPermission = true;
            pdfSecurity.ModifyDocumentPermission = false;
            pdfSecurity.ExtractContentPermission = false;

            // Assign the security options to the PDF save options.
            pdfSaveOptions.SecurityOptions = pdfSecurity;

            // Save the workbook as a PDF while preserving the protection metadata.
            workbook.Save(pdfPath, pdfSaveOptions);
        }

        // Example usage
        public static void Main()
        {
            string inputWorkbook = "ProtectedWorkbook.xlsx";
            string outputPdf = "ProtectedWorkbook.pdf";
            string workbookPassword = "password123"; // set to null if workbook is not password protected

            Convert(inputWorkbook, outputPdf, workbookPassword);

            Console.WriteLine("Conversion completed. PDF saved to: " + outputPdf);
        }
    }
}

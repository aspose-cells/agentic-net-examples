// Title: Decrypt a password‑protected PDF generated from Excel and extract its text with Aspose.Cells & Aspose.Pdf for .NET
// Description: This example shows how to load a password‑protected Excel workbook, convert it to a secured PDF, then open the PDF with the owner password, remove extraction restrictions, and programmatically retrieve all text using Aspose.Pdf APIs.
// Keywords: Aspose.Cells PDF security | Aspose.Pdf decrypt password protected PDF | extract text from secured PDF C# | Excel to PDF conversion with passwords .NET | remove PDF content restrictions Aspose | owner password PDF Aspose.Pdf | programmatic PDF text extraction C#
// Common Searches: How to open a password‑protected PDF created from Excel in C# | Aspose.Pdf remove security and read text from PDF | Convert protected Excel workbook to PDF and extract content | C# code to decrypt PDF and get plain text using Aspose | Set and later clear PDF permissions with Aspose.Cells and Aspose.Pdf
// Developer Intent: Load a protected Excel file, generate a PDF with owner/user passwords, then unlock the PDF and extract its full text using Aspose libraries.
// Use Cases: Automated pipeline that creates confidential PDFs from Excel and later reads their content for indexing or compliance checks. | Batch processing of secured reports where the PDF must be decrypted to feed downstream analytics. | Regulatory workflow that requires generation of a locked PDF for distribution, followed by programmatic verification of its textual content.
// AI Prompts: Generate C# code that opens the password‑protected PDF created above, supplies the owner password, disables security restrictions, and writes the extracted text to a .txt file using Aspose.Pdf. | Show how to catch and handle an InvalidPasswordException when the supplied PDF password is incorrect in a .NET application. | Explain the difference between user and owner passwords in PDF security and how Aspose.Pdf can modify permissions after a file has been saved.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfDecryptExample
{
    // This example shows how to load a password‑protected Excel workbook, convert it to a secured PDF, then open the PDF with the owner password, remove extraction restrictions, and programmatically retrieve all text using Aspose.Pdf APIs.
    class Program
    {
        static void Main()
        {
            try
            {
                const string excelPath = "ProtectedWorkbook.xlsx";
                const string pdfPath = "ProtectedOutput.pdf";

                // Verify that the source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                    return;
                }

                // -----------------------------------------------------------------
                // 1. Load the password‑protected Excel workbook
                // -----------------------------------------------------------------
                var loadOptions = new LoadOptions
                {
                    Password = "excelPwd" // Excel file password
                };
                var workbook = new Workbook(excelPath, loadOptions);

                // -----------------------------------------------------------------
                // 2. Save the workbook as a password‑protected PDF
                // -----------------------------------------------------------------
                var pdfSaveOptions = new PdfSaveOptions();
                var pdfSecurity = new PdfSecurityOptions
                {
                    OwnerPassword = "ownerPwd",    // Owner password (full access)
                    UserPassword = "userPwd",      // User password (restricted access)
                    ExtractContentPermission = false, // Disallow content extraction
                    PrintPermission = true,
                    ModifyDocumentPermission = false
                };
                pdfSaveOptions.SecurityOptions = pdfSecurity;

                workbook.Save(pdfPath, pdfSaveOptions);
                Console.WriteLine($"PDF saved successfully to \"{pdfPath}\" with security settings.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

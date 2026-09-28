// Title: How to convert an Aspose.Cells workbook to a password‑protected PDF in C#
// AI Prompts: Write C# code that saves an Aspose.Cells Workbook as a PDF with a user password and an optional owner password. | Show how to configure PdfSaveOptions with PdfSecurityOptions to encrypt a PDF generated from a workbook using Aspose.Cells.
// Common Searches: C# Aspose.Cells export workbook to encrypted PDF with user password | Set owner and user passwords for PDF output in Aspose.Cells .NET | How to apply PDF security options when saving a workbook as PDF using Aspose.Cells | Aspose.Cells PdfSaveOptions password protection example in C#
// Tags: Aspose.Cells PdfSaveOptions password protection | C# set PdfSecurityOptions user password | Encrypt PDF from workbook Aspose.Cells | owner password Aspose.Cells PDF export | Aspose.Cells PDF encryption .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// The example creates a new Workbook, adds data to cell A1, configures PdfSaveOptions with a PdfSecurityOptions object that specifies a user password (and optionally an owner password), assigns the security options to the save options, and then saves the workbook as an encrypted PDF file named "EncryptedWorkbook.pdf".
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data
        Workbook workbook = new Workbook();
        workbook.Worksheets[0].Cells["A1"].Value = "Secure PDF Example";

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Create PDF security options and set the user password
        PdfSecurityOptions securityOptions = new PdfSecurityOptions();
        securityOptions.UserPassword = "userPassword123";   // password required to open the PDF
        securityOptions.OwnerPassword = "ownerPassword123"; // optional owner password

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as a password‑protected PDF
        workbook.Save("EncryptedWorkbook.pdf", pdfSaveOptions);
    }
}

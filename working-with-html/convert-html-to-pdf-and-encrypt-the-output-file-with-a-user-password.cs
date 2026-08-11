// Title: Convert HTML to Password‑Protected PDF (User Password) with Aspose.Cells for .NET
// Description: Loads an HTML file into an Aspose.Cells Workbook, configures PdfSecurityOptions with a user password (and optional owner password), and saves the workbook as an encrypted PDF using PdfSaveOptions.
// Keywords: Aspose.Cells HTML to PDF | C# PDF encryption | PdfSecurityOptions user password | password protected PDF .NET | encrypt PDF from HTML | Aspose.Cells PdfSaveOptions | secure PDF generation | Aspose.Cells example C#
// Common Searches: Aspose.Cells encrypt PDF with user password | How to add password protection to PDF generated from HTML in C# | Convert HTML to PDF and set user password using Aspose.Cells | PdfSecurityOptions example C# | Save workbook as encrypted PDF Aspose.Cells | HTML to PDF encryption .NET tutorial
// Developer Intent: Create a PDF from HTML and protect it with a user‑password using Aspose.Cells.
// Use Cases: Securely share a generated report by converting an HTML template to a password‑protected PDF. | Deliver confidential invoices as encrypted PDFs generated from HTML layouts. | Archive web‑page content in a PDF that requires a password to open. | Distribute internal documentation with PDF encryption to meet compliance requirements.
// AI Prompts: Generate C# code that saves an HTML workbook as a PDF protected only by a user password with Aspose.Cells. | Show how to set 128‑bit encryption and disable printing using PdfSecurityOptions. | Explain how to open a PDF encrypted with Aspose.Cells in C# without the owner password. | Provide an example that uses only a user password and leaves the owner password blank.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

// Loads an HTML file into an Aspose.Cells Workbook, configures PdfSecurityOptions with a user password (and optional owner password), and saves the workbook as an encrypted PDF using PdfSaveOptions.
class HtmlToPdfEncrypt
{
    static void Main()
    {
        // Load the HTML file into a workbook
        Workbook workbook = new Workbook("input.html");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Configure PDF security options with a user password (and optional owner password)
        PdfSecurityOptions securityOptions = new PdfSecurityOptions
        {
            UserPassword = "UserPassword123",   // password required to open the PDF
            OwnerPassword = "OwnerPassword123" // password that grants full access
        };

        // Assign the security options to the PDF save options
        pdfSaveOptions.SecurityOptions = securityOptions;

        // Save the workbook as an encrypted PDF
        workbook.Save("encrypted_output.pdf", pdfSaveOptions);
    }
}

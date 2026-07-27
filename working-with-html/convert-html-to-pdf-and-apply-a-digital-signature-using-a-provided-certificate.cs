// Title: C# – Convert HTML to PDF/A and digitally sign with a PFX certificate using Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, creates a DigitalSignature from a PFX certificate, embeds the signature in the workbook, saves the signed workbook as XLSX, and then exports it to a PDF/A‑1b document with PdfSaveOptions.
// Keywords: Aspose.Cells | HTML to PDF conversion C# | PDF/A compliance | digital signature PFX | X509Certificate2 | LoadOptions Html | PdfSaveOptions | Workbook digital signing | C# code example | convert HTML to Excel
// Common Searches: Aspose.Cells convert HTML to PDF/A C# | How to sign an Excel workbook with a PFX certificate | Add digital signature before saving PDF with Aspose.Cells | Load HTML into Workbook and export to PDF/A | C# example for PDF/A‑1b generation with digital signature
// Developer Intent: Generate a PDF/A‑1b file from an HTML source and apply a cryptographic signature using a PFX certificate via Aspose.Cells.
// Use Cases: Create regulatory‑compliant PDF/A reports from web‑based HTML templates. | Automate batch conversion of HTML invoices to signed PDF/A archives. | Produce a signed Excel version of an HTML dashboard before distributing it as PDF.
// AI Prompts: Show how to add a visible signature image to the PDF produced by Aspose.Cells. | Provide code to select a certificate from the Windows certificate store instead of a PFX file. | Explain error handling for expired certificates or incorrect passwords during signing.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;
using Aspose.Cells.Rendering;

// Loads an HTML file into an Aspose.Cells Workbook, creates a DigitalSignature from a PFX certificate, embeds the signature in the workbook, saves the signed workbook as XLSX, and then exports it to a PDF/A‑1b document with PdfSaveOptions.
class HtmlToPdfWithSignature
{
    static void Main()
    {
        try
        {
            // Input HTML file, output PDF file, and optional signed Excel file paths
            string htmlPath = "input.html";
            string pdfPath = "output.pdf";
            string signedXlsxPath = "signed_output.xlsx";

            // Certificate file (PFX) and its password
            string certPath = "certificate.pfx";
            string certPassword = "yourPassword";

            // Verify required files exist
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException($"HTML input file not found: {htmlPath}");
            if (!File.Exists(certPath))
                throw new FileNotFoundException($"Certificate file not found: {certPath}");

            // Load the HTML file into a workbook (using LoadOptions for HTML)
            var loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Load the certificate that contains a private key
            X509Certificate2 certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);

            // Create a digital signature using the certificate
            DigitalSignature signature = new DigitalSignature(
                certificate,
                "HTML to PDF conversion",
                DateTime.UtcNow);

            // Add the signature to a collection and apply it to the workbook
            DigitalSignatureCollection signatures = new DigitalSignatureCollection();
            signatures.Add(signature);
            workbook.SetDigitalSignature(signatures);

            // Save the signed workbook (optional, keeps the signature in the Excel file)
            workbook.Save(signedXlsxPath, SaveFormat.Xlsx);

            // Prepare PDF save options (e.g., set PDF/A compliance)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b
            };

            // Convert the signed workbook to PDF
            workbook.Save(pdfPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

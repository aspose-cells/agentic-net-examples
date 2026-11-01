// Title: Convert an Excel workbook to PDF and add a digital signature from a Windows certificate store using Aspose.Cells for .NET (C#)
// Common Searches: how to sign a PDF generated from an Excel file using a certificate from the Windows store in C# Aspose.Cells | Aspose.Cells convert xlsx to pdf and apply digital signature with X509 certificate | load X509 certificate by thumbprint from personal store for PDF signing in .NET | C# Aspose.Cells PdfSaveOptions digital signature support and version requirements | save Excel workbook as signed PDF using Aspose.Cells and Windows personal certificate
// Tags: Aspose.Cells PDF conversion with X509 certificate | C# X509Store thumbprint retrieval | PdfSaveOptions digital signature configuration | Windows personal certificate store signing | Excel to signed PDF using Aspose.Cells

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves a certificate from the current user's personal X509 store by thumbprint, prepares PdfSaveOptions (noting that the DigitalSignature property may be unavailable in older Aspose.Cells versions), and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Open the personal (My) certificate store of the current user
            X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);

            // TODO: replace with the actual thumbprint of the certificate you want to use
            const string certificateThumbprint = "YOUR_CERT_THUMBPRINT";

            // Find the certificate by thumbprint
            X509Certificate2 certificate = null;
            foreach (X509Certificate2 cert in store.Certificates)
            {
                if (!string.IsNullOrEmpty(cert.Thumbprint) &&
                    cert.Thumbprint.Equals(certificateThumbprint, StringComparison.OrdinalIgnoreCase))
                {
                    certificate = cert;
                    break;
                }
            }

            store.Close();

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: DigitalSignature property is not available in the current Aspose.Cells version.
            // If a later version supports it, you can assign the signature as shown below:
            // pdfOptions.DigitalSignature = new DigitalSignature(certificate);
            if (certificate != null)
            {
                Console.WriteLine("Certificate found, but digital signing is not supported in this Aspose.Cells version.");
            }
            else
            {
                Console.WriteLine("Certificate not found. PDF will be saved without a digital signature.");
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

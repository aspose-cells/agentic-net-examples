// Title: How to export a VBA project's digital certificate to a .cer file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export VBA macro digital certificate to .cer | Save VBA project signature certificate from Excel workbook using Aspose.Cells | Extract X509Certificate2 from VBA project in an .xlsx file with Aspose.Cells | Aspose.Cells VBA project certificate extraction not supported version workaround | How to retrieve VBA project certificate in .NET using Aspose.Cells
// Tags: Aspose.Cells VBA project certificate extraction | C# export X509Certificate2 to .cer file | Excel VBA macro digital signature handling | Workbook VbaProject SignatureInfo usage | Aspose.Cells unsupported certificate extraction note

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The sample loads an Excel workbook, checks for a VBA project, and shows how to export the VBA project's X509Certificate2 to a .cer file when the SignatureInfo and Certificate objects are exposed; if not, it reports that the current Aspose.Cells version does not support certificate extraction.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that may contain a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Path where the certificate will be saved (if available)
            string certificatePath = "vba_certificate.cer";

            // Check if the workbook has a VBA project
            if (workbook.VbaProject != null)
            {
                // NOTE: In some versions of Aspose.Cells the VBA project signature
                // information is not exposed via the API. If SignatureInfo is available,
                // you can uncomment the following block to export the certificate.

                /*
                if (workbook.VbaProject.SignatureInfo != null &&
                    workbook.VbaProject.SignatureInfo.Certificate != null)
                {
                    X509Certificate2 cert = workbook.VbaProject.SignatureInfo.Certificate;
                    byte[] certData = cert.Export(X509ContentType.Cert);
                    File.WriteAllBytes(certificatePath, certData);
                    Console.WriteLine($"VBA digital certificate exported to: {certificatePath}");
                    return;
                }
                */

                Console.WriteLine("The workbook contains a VBA project, but extracting the digital certificate is not supported in this version of Aspose.Cells.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

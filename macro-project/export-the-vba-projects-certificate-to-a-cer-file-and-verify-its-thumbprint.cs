// Title: Export a VBA project's digital certificate to a .cer file and verify its thumbprint using Aspose.Cells for .NET (C#)
// Common Searches: C# export VBA project certificate to .cer using Aspose.Cells | How to read and verify VBA macro certificate thumbprint in .NET | Aspose.Cells extract VBA project digital signature and compare thumbprint | Save VBA project X509 certificate as .cer file with C# | Validate exported VBA macro certificate thumbprint with X509Certificate2
// Tags: export VBA certificate Aspose.Cells C# | save X509 certificate as .cer file | verify certificate thumbprint .NET | access VBA project digital signature Aspose | compare X509Certificate2 thumbprints

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an Excel workbook, accesses its VBA project's digital signature, exports the attached X509 certificate to a .cer file, writes the file, retrieves the certificate's thumbprint, reloads the .cer file as an X509Certificate2 object, and compares both thumbprints to confirm they match.
class VbaCertificateExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "vba_cert.cer";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains a VBA project
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project and its digital signature using dynamic to avoid compile‑time binding issues
            dynamic vbaProject = workbook.VbaProject;
            dynamic signature = vbaProject?.Signature;
            dynamic certificate = signature?.Certificate;

            if (certificate == null)
            {
                Console.WriteLine("No certificate is attached to the VBA project.");
                return;
            }

            // Export the certificate to a .cer file
            byte[] certBytes = certificate.Export(X509ContentType.Cert);
            File.WriteAllBytes(outputPath, certBytes);
            Console.WriteLine($"Certificate exported to {outputPath}");

            // Verify the thumbprint of the exported certificate
            string thumbprint = certificate.Thumbprint;
            Console.WriteLine($"Certificate Thumbprint: {thumbprint}");

            // Load the exported .cer file and compare thumbprints
            X509Certificate2 loadedCert = new X509Certificate2(File.ReadAllBytes(outputPath));
            if (string.Equals(thumbprint, loadedCert.Thumbprint, StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Thumbprint verification succeeded.");
            }
            else
            {
                Console.WriteLine("Thumbprint verification failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

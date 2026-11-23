// Title: Export a VBA project's X509 certificate to a .cer file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export VBA project certificate to .cer file | How to retrieve X509Certificate2 from an Excel VBA macro with Aspose.Cells | Save VBA macro signature certificate as a DER encoded file in .NET | Extract VBA project digital signature using Aspose.Cells C# example | Write VBA project certificate bytes to disk with Aspose.Cells
// Tags: Aspose.Cells VBA certificate extraction | C# extract X509Certificate2 from Excel VBA | write DER encoded certificate to file | retrieve VBA project signature via reflection | save VBA macro certificate as .cer

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject

// The sample loads an Excel workbook, accesses its VBA project, uses reflection to obtain the Signature and Certificate objects, exports the X509Certificate2 as a DER‑encoded byte array, and writes the resulting .cer file to disk.
class ExportVbaCertificate
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "VbaCertificate.cer";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the VBA project
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Use reflection to obtain the Signature property (may not exist in older versions)
            object signatureObj = vbaProject.GetType().GetProperty("Signature")?.GetValue(vbaProject);
            if (signatureObj == null)
            {
                Console.WriteLine("The VBA project does not contain a signature.");
                return;
            }

            // Use reflection to obtain the Certificate property from the signature
            object certObj = signatureObj.GetType().GetProperty("Certificate")?.GetValue(signatureObj);
            if (certObj is not X509Certificate2 certificate || certificate == null)
            {
                Console.WriteLine("The VBA signature does not contain a valid X509 certificate.");
                return;
            }

            // Export the certificate to a byte array (DER encoded)
            byte[] certBytes = certificate.Export(X509ContentType.Cert);

            // Write the certificate bytes directly to the output file
            File.WriteAllBytes(outputPath, certBytes);

            Console.WriteLine($"VBA project certificate exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

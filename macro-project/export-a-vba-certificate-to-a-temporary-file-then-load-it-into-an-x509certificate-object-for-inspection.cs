// Title: Export a VBA project's signing certificate to a temporary .cer file and examine it with X509Certificate2 using Aspose.Cells for .NET
// AI Prompts: Extract the VBA project's signing certificate from an Excel workbook, save it to a temporary .cer file, and load it into an X509Certificate2 object for detailed inspection. | Enhance the sample to display additional certificate fields such as serial number, public key, and signature algorithm after exporting the VBA certificate. | Refactor the code to accept a Stream containing the workbook instead of a file path while preserving the export and inspection of the VBA signing certificate.
// Common Searches: how to use Aspose.Cells to export a VBA signing certificate from an .xlsm file | C# read VBA project certificate and get subject issuer thumbprint | temporary file handling when exporting VBA certificate with Aspose.Cells | load exported .cer file into X509Certificate2 for inspection in .NET
// Tags: Aspose.Cells export VBA certificate to .cer | X509Certificate2 load certificate from file | VbaSignatureInfo ExportCertificate usage | temporary file cleanup after certificate export | inspect VBA project signing certificate .NET

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaSignatureInfo (assembly reference may be needed)

// The program loads an Excel workbook, checks for a VBA project, uses VbaSignatureInfo.ExportCertificate to write the signing certificate to a temporary .cer file, reads the file into an X509Certificate2 object, prints key certificate details (subject, issuer, thumbprint, validity period), and finally deletes the temporary file.
class VbaCertificateInspector
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook that contains a VBA project with a certificate.
            Workbook workbook = new Workbook(inputPath);

            // Ensure that the workbook actually has a VBA project.
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Use dynamic to access SignatureInfo (may not be present in older library versions).
            dynamic vbaProject = workbook.VbaProject;
            dynamic signatureInfo = null;

            try
            {
                signatureInfo = vbaProject.SignatureInfo;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // Property not available in the current Aspose.Cells version.
            }

            if (signatureInfo == null)
            {
                Console.WriteLine("The VBA project does not contain a signing certificate or the API is unavailable.");
                return;
            }

            // Create a temporary file to hold the exported certificate.
            string tempCertPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cer");

            try
            {
                // Export the VBA project's signing certificate to the temporary file.
                signatureInfo.ExportCertificate(tempCertPath);

                // Load the exported certificate into an X509Certificate2 object for inspection.
                byte[] certBytes = File.ReadAllBytes(tempCertPath);
                X509Certificate2 certificate = new X509Certificate2(certBytes);

                // Output useful information about the certificate.
                Console.WriteLine($"Certificate Subject: {certificate.Subject}");
                Console.WriteLine($"Certificate Issuer: {certificate.Issuer}");
                Console.WriteLine($"Thumbprint: {certificate.Thumbprint}");
                Console.WriteLine($"Valid From: {certificate.NotBefore}");
                Console.WriteLine($"Valid To: {certificate.NotAfter}");
            }
            finally
            {
                // Clean up the temporary certificate file.
                if (File.Exists(tempCertPath))
                {
                    File.Delete(tempCertPath);
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

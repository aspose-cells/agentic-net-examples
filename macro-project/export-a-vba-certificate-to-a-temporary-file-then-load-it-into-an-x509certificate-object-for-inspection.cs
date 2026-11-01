// Title: Export a VBA project's signing certificate to a temporary .cer file and examine it with X509Certificate2 using Aspose.Cells in C#
// Common Searches: how to extract a VBA macro signing certificate from an .xlsm workbook with Aspose.Cells C# | C# code to export a VBA project's certificate to a temporary file and read it with X509Certificate2 | using Aspose.Cells to retrieve and inspect the digital signature of a VBA project | load exported VBA certificate into X509Certificate2 to view subject and thumbprint | save VBA project signature as .cer then delete the file in C#
// Tags: Aspose.Cells export VBA certificate to .cer | C# temporary certificate file handling | X509Certificate2 load from byte array | inspect VBA project digital signature details | cleanup temporary .cer after inspection

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;

// The example loads an .xlsm workbook with Aspose.Cells, checks for a VBA project, accesses its Signature via dynamic binding, exports the signing certificate to a temporary .cer file, reads the file into an X509Certificate2 object, prints key certificate properties (subject, issuer, validity period, thumbprint), and finally removes the temporary file.
class VbaCertificateInspector
{
    static void Main()
    {
        // Path to the Excel file containing VBA project
        string excelPath = @"C:\Path\To\YourWorkbook.xlsm";

        // Verify the workbook file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Workbook file not found: {excelPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Ensure the workbook has a VBA project
            if (workbook.VbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Use dynamic to access the Signature property (avoids compile‑time binding issues)
            dynamic vbaProject = workbook.VbaProject;
            if (vbaProject == null)
            {
                Console.WriteLine("VBA project is unavailable.");
                return;
            }

            // Attempt to retrieve the signature object
            dynamic signature = null;
            try
            {
                signature = vbaProject.Signature;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // Property not found – the workbook is likely unsigned
            }

            if (signature == null)
            {
                Console.WriteLine("The VBA project is not signed with a certificate.");
                return;
            }

            // Create a temporary file for the exported certificate
            string tempCertPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cer");

            try
            {
                // Export the VBA project's signing certificate to the temporary file
                signature.ExportCertificate(tempCertPath);

                // Load the certificate for inspection (avoid obsolete constructor)
                X509Certificate2 certificate = new X509Certificate2(File.ReadAllBytes(tempCertPath));

                // Output useful information about the certificate
                Console.WriteLine($"Certificate Subject: {certificate.Subject}");
                Console.WriteLine($"Certificate Issuer: {certificate.Issuer}");
                Console.WriteLine($"Valid From: {certificate.NotBefore}");
                Console.WriteLine($"Valid To: {certificate.NotAfter}");
                Console.WriteLine($"Thumbprint: {certificate.Thumbprint}");
            }
            finally
            {
                // Clean up the temporary certificate file
                if (File.Exists(tempCertPath))
                {
                    try
                    {
                        File.Delete(tempCertPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

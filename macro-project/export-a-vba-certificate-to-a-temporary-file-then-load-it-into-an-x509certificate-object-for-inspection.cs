// Title: Export a VBA project's signing certificate to a temporary .cer file and load it into X509Certificate2 with Aspose.Cells for .NET
// Common Searches: C# export VBA project signing certificate from .xlsm using Aspose.Cells | how to load an exported VBA certificate into X509Certificate2 in .NET | Aspose.Cells ExportCertificate method example for retrieving VBA certificate | read VBA macro signing certificate details with C# and Aspose.Cells | using reflection to call ExportCertificate on a VBA project in Aspose.Cells
// Tags: Aspose.Cells export VBA certificate | load X509Certificate2 from .cer file C# | inspect VBA signing certificate .NET | reflection invoke ExportCertificate method | temporary certificate file cleanup C# | read VBA project certificate Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Security.Cryptography.X509Certificates;

// The sample loads an .xlsm workbook, checks for a VBA project, uses reflection to call the ExportCertificate method and writes the signing certificate to a temporary .cer file, reads the file into an X509Certificate2 object, prints certificate details (subject, issuer, thumbprint, validity dates), and finally deletes the temporary file while handling possible errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Ensure the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Verify that the workbook contains a VBA project
        if (workbook.VbaProject == null)
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
            return;
        }

        try
        {
            // Attempt to locate the ExportCertificate method via reflection (available in newer versions)
            var exportMethod = workbook.VbaProject.GetType().GetMethod(
                "ExportCertificate",
                new[] { typeof(string) });

            if (exportMethod != null)
            {
                // Export the VBA project's certificate to a temporary file
                string tempCertPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cer");
                exportMethod.Invoke(workbook.VbaProject, new object[] { tempCertPath });

                // Load the exported certificate (using byte[] ctor to avoid obsolete API)
                var certBytes = File.ReadAllBytes(tempCertPath);
                var certificate = new X509Certificate2(certBytes);

                // Display certificate information
                Console.WriteLine($"Subject: {certificate.Subject}");
                Console.WriteLine($"Issuer: {certificate.Issuer}");
                Console.WriteLine($"Thumbprint: {certificate.Thumbprint}");
                Console.WriteLine($"Valid From: {certificate.NotBefore}");
                Console.WriteLine($"Valid To: {certificate.NotAfter}");

                // Clean up the temporary certificate file
                File.Delete(tempCertPath);
            }
            else
            {
                Console.WriteLine("ExportCertificate method is not available in this Aspose.Cells version.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the VBA certificate: {ex.Message}");
        }
    }
}

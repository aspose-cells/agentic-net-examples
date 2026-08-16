// Title: Export VBA Project Certificate to .cer and Verify Thumbprint with Aspose.Cells for .NET
// Description: Shows how to load a macro‑enabled workbook, detect a signed VBA project, extract its CertRawData, save it as a .cer file, load the certificate with X509Certificate2, and compare thumbprints to confirm integrity.
// Keywords: Aspose.Cells | VBA certificate export | .cer file | thumbprint verification | VbaProject CertRawData | signed macro workbook | X509Certificate2 | C# example | security audit
// Common Searches: Aspose.Cells export VBA certificate | How to get certificate from signed VBA project .xlsm | Save VBA project certificate as .cer file C# | Compare VBA certificate thumbprint with Aspose.Cells | Check if workbook has signed VBA macro using Aspose.Cells
// Developer Intent: Extract the certificate from a signed VBA project, write it to a .cer file, and verify that its thumbprint matches the original certificate.
// Use Cases: Archive the certificate from a signed macro workbook for compliance records. | Perform a security audit by reading the thumbprint with X509Certificate2. | Validate that the exported .cer file is identical to the in‑memory certificate. | Detect workbooks that lack a signed VBA project and handle them gracefully.
// AI Prompts: Generate C# code using Aspose.Cells to export the CertRawData of a signed VBA project to a .cer file. | Show how to load a .cer file with X509Certificate2 and compare its thumbprint to the original VBA certificate. | Explain error handling when a workbook does not contain a signed VBA project in Aspose.Cells. | Create a PowerShell snippet that verifies the thumbprint of a VBA certificate exported by Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Security.Cryptography.X509Certificates;

namespace AsposeCellsExamples
{
    // Shows how to load a macro‑enabled workbook, detect a signed VBA project, extract its CertRawData, save it as a .cer file, load the certificate with X509Certificate2, and compare thumbprints to confirm integrity.
    class ExportVbaCertificate
    {
        static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Path where the extracted certificate will be saved
            string certificatePath = "VbaCertificate.cer";

            // Verify workbook file exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Workbook file not found: {signedWorkbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Get the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check that the VBA project is signed and that certificate data exists
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Export the raw certificate bytes to a .cer file
                File.WriteAllBytes(certificatePath, vbaProject.CertRawData);
                Console.WriteLine($"Certificate exported to: {certificatePath}");

                // Load the exported certificate from file
                X509Certificate2 exportedCertificate = new X509Certificate2();
                exportedCertificate.Import(certificatePath);
                Console.WriteLine($"Exported certificate thumbprint: {exportedCertificate.Thumbprint}");

                // Load the original certificate directly from the raw data for verification
                X509Certificate2 originalCertificate = new X509Certificate2();
                originalCertificate.Import(vbaProject.CertRawData);
                Console.WriteLine($"Original certificate thumbprint: {originalCertificate.Thumbprint}");

                // Verify that the thumbprints match
                bool thumbprintsMatch = string.Equals(
                    exportedCertificate.Thumbprint,
                    originalCertificate.Thumbprint,
                    StringComparison.OrdinalIgnoreCase);

                Console.WriteLine($"Thumbprint verification result: {thumbprintsMatch}");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
            }
        }
    }
}

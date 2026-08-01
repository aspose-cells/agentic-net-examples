// Title: Export a Signed VBA Project Certificate to a Password‑Protected PFX File with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled .xlsm workbook, accesses its VbaProject, verifies the project is signed, imports the embedded X509 certificate, exports it as a password‑protected PFX byte array, and writes the .pfx file to disk using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# VBA certificate export | export VBA project certificate | password protected PFX | VbaProject CertRawData | macro-enabled workbook | extract signed VBA certificate | X509Certificate2 Aspose.Cells
// Common Searches: export VBA project certificate to pfx Aspose.Cells | C# extract signed macro workbook certificate | save VBA signing certificate as password protected PFX | Aspose.Cells retrieve VBA project CertRawData | how to backup VBA project certificate .NET
// Developer Intent: Create a password‑protected PFX file from the signing certificate embedded in a signed VBA project.
// Use Cases: Archive the signing certificate of a macro‑enabled workbook for compliance audits. | Migrate a VBA project's signing certificate to another system while keeping the private key encrypted. | Back up the VBA project's certificate before re‑signing or updating the macro code.
// AI Prompts: Write C# code that loads a .xlsm file with Aspose.Cells, extracts the VBA project's certificate, and saves it as a password‑protected .pfx file. | Explain how to handle an unsigned VBA project or missing certificate data when using Aspose.Cells. | Provide a step‑by‑step guide to verify that the exported PFX contains the private key using .NET security APIs.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled .xlsm workbook, accesses its VbaProject, verifies the project is signed, imports the embedded X509 certificate, exports it as a password‑protected PFX byte array, and writes the .pfx file to disk using Aspose.Cells for .NET.
    public class ExportVbaCertificateToPfx
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Verify that the workbook file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed and certificate data is available
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Load the certificate from raw data (avoids obsolete constructor)
                X509Certificate2 certificate = new X509Certificate2();
                certificate.Import(vbaProject.CertRawData);

                // Define a password to protect the exported .pfx file
                string pfxPassword = "ExportPassword123";

                // Export the certificate (including private key if present) to a PFX byte array
                byte[] pfxData = certificate.Export(X509ContentType.Pfx, pfxPassword);

                // Write the PFX data to a file for secure storage
                string outputPfxPath = "VbaProjectCertificate.pfx";
                File.WriteAllBytes(outputPfxPath, pfxData);

                Console.WriteLine($"Certificate exported successfully to '{outputPfxPath}'.");
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
            }

            // No need to save the workbook for this operation; the focus is on exporting the certificate.
        }
    }
}

// Title: Export a signed VBA project's certificate to a password‑protected PFX file using Aspose.Cells for .NET
// Description: Loads a signed .xlsm workbook with Aspose.Cells, accesses its VbaProject, verifies the presence of a signing certificate, imports the raw certificate data into an X509Certificate2 object, and exports it as a password‑protected .pfx file saved to disk.
// Keywords: Aspose.Cells VBA certificate export | C# export VBA project certificate | password protected PFX .NET | X509Certificate2 Import CertRawData | VbaProject IsSigned check | Excel macro signing backup
// Common Searches: export VBA project certificate Aspose.Cells C# | save signed macro certificate as .pfx | how to protect exported certificate with password .NET | retrieve CertRawData from .xlsm workbook | backup VBA signing certificate programmatically
// Developer Intent: Extract the signing certificate from a signed VBA project and save it as a password‑protected .pfx file for secure archival or migration.
// Use Cases: Create a secure backup of a VBA project's signing certificate for compliance audits. | Migrate a macro's signing certificate to another server or development environment. | Automate verification that a VBA project is signed before performing certificate export.
// AI Prompts: Write C# code that loads a signed .xlsm file with Aspose.Cells, checks VbaProject.IsSigned, and exports the certificate to a password‑protected PFX. | Explain how to use X509Certificate2.Import with VbaProject.CertRawData and then export the certificate as a PFX using a custom password. | Suggest robust error‑handling patterns for exporting a VBA project's certificate to a .pfx file in a .NET application.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm workbook with Aspose.Cells, accesses its VbaProject, verifies the presence of a signing certificate, imports the raw certificate data into an X509Certificate2 object, and exports it as a password‑protected .pfx file saved to disk.
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file '{workbookPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed and certificate data is available
                if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
                {
                    // Load the certificate from the raw data using Import (avoids obsolete ctor)
                    X509Certificate2 certificate = new X509Certificate2();
                    certificate.Import(vbaProject.CertRawData);

                    // Define a password to protect the exported .pfx file
                    string exportPassword = "StrongPassword123";

                    // Export the certificate (including private key if present) to a PFX byte array
                    byte[] pfxData = certificate.Export(X509ContentType.Pfx, exportPassword);

                    // Save the PFX data to a file
                    string outputPath = "VbaProjectCertificate.pfx";
                    File.WriteAllBytes(outputPath, pfxData);

                    Console.WriteLine($"Certificate exported successfully to '{outputPath}'.");
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed or does not contain certificate data.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}

// Title: Export a VBA Project's Digital Certificate to a File Stream with Aspose.Cells (.NET)
// Description: Demonstrates how to load a workbook that contains a signed VBA project, verify the signature, extract the raw certificate bytes via `VbaProject.CertRawData`, and write them to a .cer file using a `FileStream` for backup or compliance purposes.
// Keywords: Aspose.Cells VBA certificate export | C# extract VbaProject CertRawData | backup signed VBA macro certificate | write certificate to file stream .NET | VbaProject.IsSigned check | digital certificate backup Aspose | export VBA project certificate C#
// Common Searches: How to export a VBA project's digital certificate with Aspose.Cells | C# code to save VBA macro certificate to .cer file | Retrieve CertRawData from VbaProject in .NET | Backup signed VBA project certificate programmatically | Aspose.Cells example for exporting VBA certificate
// Developer Intent: Extract and save the digital certificate of a signed VBA project for backup or audit.
// Use Cases: Create a secure backup of a signed macro's certificate before editing the workbook. | Store the certificate for compliance audits or trust verification across systems. | Automate migration of VBA project certificates to a central repository.
// AI Prompts: Generate C# code using Aspose.Cells that checks if a workbook's VBA project is signed and writes its certificate to a .cer file. | Provide a reusable method that returns the certificate bytes from a VbaProject, handling missing files and empty data gracefully. | Create error‑handling and logging logic for exporting a VBA project's digital certificate to a file stream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Demonstrates how to load a workbook that contains a signed VBA project, verify the signature, extract the raw certificate bytes via `VbaProject.CertRawData`, and write them to a .cer file using a `FileStream` for backup or compliance purposes.
    public class VbaProjectCertificateExportDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An unexpected error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists
                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate to a file for backup
                    string certPath = "VbaCertificateBackup.cer";
                    using (FileStream stream = new FileStream(certPath, FileMode.Create, FileAccess.Write))
                    {
                        stream.Write(certData, 0, certData.Length);
                    }

                    Console.WriteLine($"Certificate exported successfully to '{certPath}'. Size: {certData.Length} bytes.");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed. No certificate to export.");
            }
        }
    }
}

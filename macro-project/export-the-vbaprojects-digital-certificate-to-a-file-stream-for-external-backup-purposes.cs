// Title: Export a VBA project's digital certificate to a file stream using Aspose.Cells for .NET
// Description: Loads a workbook with a signed VBA project, verifies the signature, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file with FileStream, and optionally loads the bytes into a MemoryStream for further processing. Includes error handling for missing files and unsigned projects.
// Keywords: Aspose.Cells | C# | VbaProject | Export VBA certificate | CertRawData | digital certificate backup | file stream | memory stream | signed macro extraction | .cer file generation
// Common Searches: how to export VBA project certificate with Aspose.Cells | save signed macro certificate to .cer file C# | Aspose.Cells extract CertRawData from VbaProject | backup VBA digital certificate using .NET | write VBA certificate bytes to file stream
// Developer Intent: Extract and save the digital certificate of a signed VBA project for backup or further validation.
// Use Cases: Create an offline backup of a signed macro's certificate before editing the workbook. | Transfer the extracted certificate to another environment for code‑signing verification. | Load the certificate into a MemoryStream to integrate with custom cryptographic checks.
// AI Prompts: Write C# code that uses Aspose.Cells to read a signed workbook, retrieve VbaProject.CertRawData, and save it as a .cer file with comprehensive error handling. | Show an example that logs each step while exporting a VBA project's certificate and writes the data to both a FileStream and a MemoryStream. | Demonstrate how to feed the exported VBA certificate into a .NET digital signature verification routine after extracting it with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a workbook with a signed VBA project, verifies the signature, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file with FileStream, and optionally loads the bytes into a MemoryStream for further processing. Includes error handling for missing files and unsigned projects.
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string workbookPath = "SignedWorkbook.xlsm";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Workbook file not found: {workbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    if (certData != null && certData.Length > 0)
                    {
                        string certPath = "VbaCertificateBackup.cer";

                        // Export the certificate to a file
                        using (FileStream fileStream = new FileStream(certPath, FileMode.Create, FileAccess.Write))
                        {
                            fileStream.Write(certData, 0, certData.Length);
                        }

                        // Optional: keep the data in a memory stream for further processing
                        using (MemoryStream memoryStream = new MemoryStream(certData))
                        {
                            // memoryStream can be used as needed
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
                    Console.WriteLine("VBA project is not signed or not present. No certificate to export.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}

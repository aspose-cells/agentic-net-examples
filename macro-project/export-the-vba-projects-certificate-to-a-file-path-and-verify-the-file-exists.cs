// Title: Export VBA Project Certificate to .cer and Verify File Creation with Aspose.Cells for .NET
// Description: Loads a signed .xlsm workbook, checks the VBA project's IsSigned flag, extracts the certificate via VbaProject.CertRawData, writes it to a .cer file, and confirms the file exists using File.Exists. Demonstrates error handling and verification steps with Aspose.Cells for .NET.
// Keywords: Aspose.Cells export VBA certificate | C# VBA project certificate extraction | VbaProject CertRawData | save .cer file from signed workbook | verify exported certificate file | Aspose.Cells VBA signing | export signed VBA project | C# file existence check
// Common Searches: How to export a VBA project's certificate with Aspose.Cells .NET | Save VBA signing certificate to .cer file in C# | Aspose.Cells retrieve VBA certificate raw data | Verify exported certificate file exists C# | Export signed VBA project certificate using Aspose.Cells
// Developer Intent: Extract a signed VBA project's certificate to a .cer file and confirm the file was created.
// Use Cases: Backup or distribute the signing certificate of a VBA-enabled workbook. | Automate compliance audits by exporting certificates from multiple signed workbooks. | Validate that a workbook's VBA project is signed before performing further processing.
// AI Prompts: Generate C# code using Aspose.Cells to export a VBA project's certificate to a specified path and verify the file exists. | Add robust error handling and logging to the VBA certificate export example. | Create a script that scans a folder of .xlsm files, exports each signed VBA project's certificate, and records success or failure.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    // Loads a signed .xlsm workbook, checks the VBA project's IsSigned flag, extracts the certificate via VbaProject.CertRawData, writes it to a .cer file, and confirms the file exists using File.Exists. Demonstrates error handling and verification steps with Aspose.Cells for .NET.
    public class ExportVbaCertificate
    {
        // Entry point for the console application
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

            // Verify the workbook file exists before proceeding
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found at path: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Check whether the VBA project is signed
                if (!vbaProject.IsSigned)
                {
                    Console.WriteLine("The VBA project is not signed. No certificate to export.");
                    return;
                }

                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data is available
                if (certData == null || certData.Length == 0)
                {
                    Console.WriteLine("Certificate raw data is empty.");
                    return;
                }

                // Define the output file path for the exported certificate
                string certFilePath = "VbaCertificate.cer";

                // Write the certificate data to the file
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate exported to: {certFilePath}");

                // Verify that the file now exists on disk
                bool fileExists = File.Exists(certFilePath);
                Console.WriteLine($"Verification - file exists: {fileExists}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

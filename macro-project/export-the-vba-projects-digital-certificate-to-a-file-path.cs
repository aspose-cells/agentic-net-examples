// Title: Export a VBA Project's Digital Certificate to a .cer File with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, accesses its VbaProject, verifies the project is signed, extracts the certificate bytes via CertRawData, and writes them to a .cer file with robust file‑existence and I/O error handling.
// Keywords: Aspose.Cells export VBA certificate | C# extract VBA project certificate | VbaProject CertRawData | save signed VBA certificate .cer | macro workbook digital signature extraction
// Common Searches: how to export VBA certificate using Aspose.Cells | retrieve signed macro certificate C# | save VBA project CertRawData to file | Aspose.Cells get VBA digital signature
// Developer Intent: Save the signing certificate of a signed VBA project from an Excel file to a .cer file.
// Use Cases: Validate macro authenticity by comparing exported certificates across workbooks. | Archive VBA signing certificates for compliance and audit trails. | Automate detection of changed macro signatures in batch‑processed Excel files.
// AI Prompts: Generate C# code that uses Aspose.Cells to read CertRawData from a signed VBA project and write it to a .cer file with error handling. | Explain the steps to confirm a VBA project is signed before exporting its certificate with Aspose.Cells for .NET. | Suggest best practices for handling file‑system exceptions when saving an exported VBA certificate.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;   // Required for VbaProject

namespace AsposeCellsExamples
{
    // Loads a macro‑enabled workbook, accesses its VbaProject, verifies the project is signed, extracts the certificate bytes via CertRawData, and writes them to a .cer file with robust file‑existence and I/O error handling.
    class ExportVbaCertificate
    {
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Destination path for the exported certificate file
            string certificatePath = "VbaCertificate.cer";

            // Verify the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed
                if (vbaProject != null && vbaProject.IsSigned)
                {
                    // Retrieve the raw certificate data
                    byte[] certData = vbaProject.CertRawData;

                    // Ensure certificate data exists before writing to file
                    if (certData != null && certData.Length > 0)
                    {
                        try
                        {
                            // Export the certificate to the specified file path
                            File.WriteAllBytes(certificatePath, certData);
                            Console.WriteLine($"Certificate exported successfully to: {certificatePath}");
                        }
                        catch (Exception ioEx)
                        {
                            Console.WriteLine($"Failed to write certificate file: {ioEx.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Certificate data is empty; nothing to export.");
                    }
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed; no certificate to export.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}

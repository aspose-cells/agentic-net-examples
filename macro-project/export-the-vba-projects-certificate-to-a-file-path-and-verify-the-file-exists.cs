// Title: Export VBA Project Certificate to .cer and Verify File with Aspose.Cells for .NET (C#)
// Description: Loads a macro-enabled workbook, checks if its VBA project is signed, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file, and confirms the file exists using File.Exists.
// Keywords: Aspose.Cells | C# | VBA certificate export | VbaProject CertRawData | export .cer file | verify certificate file | signed VBA project | macro workbook | file existence check
// Common Searches: export signed VBA project certificate Aspose.Cells C# | how to save VBA certificate to .cer file using Aspose | check if VBA project is signed and write certificate to disk | verify exported VBA certificate file exists | retrieve VbaProject CertRawData Aspose.Cells
// Developer Intent: Export the certificate of a signed VBA project to a .cer file and confirm that the file was successfully created.
// Use Cases: Archive the certificate of a signed macro workbook for compliance audits. | Validate that a workbook’s VBA code is signed before processing it further. | Store the certificate locally to use in digital signature verification workflows.
// AI Prompts: Generate C# code with Aspose.Cells that extracts a VBA project's certificate and saves it as a .cer file, then checks the file’s existence. | Create a reusable method that returns true only when a workbook’s VBA project is signed and its certificate is written to a specified path. | Write robust error‑handling for scenarios where the VBA project is unsigned, the certificate data is empty, or the file cannot be written.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// Loads a macro-enabled workbook, checks if its VBA project is signed, extracts the raw certificate via VbaProject.CertRawData, writes it to a .cer file, and confirms the file exists using File.Exists.
class ExportVbaCertificate
{
    static void Main()
    {
        // Path to the workbook that contains a signed VBA project
        string workbookPath = "SignedWorkbook.xlsm";

        // Destination path for the exported certificate file
        string certificatePath = "VbaCertificate.cer";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Verify that the VBA project is signed
        if (vbaProject.IsSigned)
        {
            // Retrieve the raw certificate data (CertRawData property)
            byte[] certData = vbaProject.CertRawData;

            // Ensure we have data before writing
            if (certData != null && certData.Length > 0)
            {
                // Export the certificate to a file
                File.WriteAllBytes(certificatePath, certData);
                Console.WriteLine($"Certificate exported to: {certificatePath}");

                // Verify that the file now exists
                bool fileExists = File.Exists(certificatePath);
                Console.WriteLine($"File exists: {fileExists}");
            }
            else
            {
                Console.WriteLine("Certificate data is empty.");
            }
        }
        else
        {
            Console.WriteLine("VBA project is not signed; no certificate to export.");
        }
    }
}

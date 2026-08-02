// Title: Export VBA Project Signing Certificate to .cer using Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, accesses its VbaProject, confirms the project is signed, extracts the raw certificate bytes via CertRawData, writes them to a MemoryStream, and copies the stream to a FileStream to create a .cer file on disk.
// Keywords: Aspose.Cells | C# | VbaProject | Export certificate | CertRawData | MemoryStream | FileStream | macro-enabled workbook | signed VBA project | extract certificate | save .cer file
// Common Searches: How to export a VBA project's signing certificate with Aspose.Cells | Aspose.Cells C# extract VBA certificate to .cer | Get CertRawData from signed macro workbook .NET | Save VBA signing certificate to file using Aspose.Cells | Export VBA certificate memory stream C#
// Developer Intent: Extract and save the signing certificate of a signed VBA project to a .cer file.
// Use Cases: Verify macro authenticity by comparing the extracted certificate with a trusted store. | Archive VBA project certificates for compliance or audit purposes. | Migrate certificates to another environment for re‑signing macros. | Integrate certificate extraction into automated workbook processing pipelines.
// AI Prompts: Write C# code that loads a .xlsm file with Aspose.Cells, checks if the VBA project is signed, and writes the CertRawData to a .cer file using MemoryStream. | Provide error‑handling examples for when a VBA project is unsigned or the certificate data is missing. | Show how to convert the extracted certificate bytes to a Base64 string and display it. | Explain how to add the exported .cer file to the Windows certificate store programmatically.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    // Loads a macro‑enabled workbook, accesses its VbaProject, confirms the project is signed, extracts the raw certificate bytes via CertRawData, writes them to a MemoryStream, and copies the stream to a FileStream to create a .cer file on disk.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the VBA project is signed before attempting to export the certificate
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data (byte array)
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate data to a MemoryStream
                    using (MemoryStream certStream = new MemoryStream(certData))
                    {
                        // Define the output file path for the certificate
                        string outputCertPath = "VbaCertificate.cer";

                        // Write the contents of the MemoryStream to disk
                        using (FileStream fileStream = new FileStream(outputCertPath, FileMode.Create, FileAccess.Write))
                        {
                            certStream.CopyTo(fileStream);
                        }

                        Console.WriteLine($"Certificate exported successfully to '{outputCertPath}'.");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or does not exist.");
            }

            // Optionally, save the workbook if any modifications were made (uses the provided save rule)
            // workbook.Save("ModifiedWorkbook.xlsm");
        }
    }
}

// Title: Export a VBA Project Certificate to a File with Aspose.Cells for .NET (C#)
// Description: Loads a macro‑enabled workbook, accesses its VbaProject, confirms the project is signed, extracts the raw certificate bytes into a MemoryStream, and writes the stream to a .cer file on disk using Aspose.Cells.
// Keywords: Aspose.Cells export VBA certificate | C# extract VBA signing certificate | VbaProject CertRawData | save .cer file from VBA project | macro‑enabled workbook certificate extraction | MemoryStream to file .NET | signed VBA project handling | Aspose.Cells VBA project certificate
// Common Searches: how to extract VBA project certificate using Aspose.Cells | save signed VBA certificate to disk C# | export VBA signing certificate to .cer file | Aspose.Cells get VbaProject certificate bytes | write MemoryStream to file in .NET
// Developer Intent: Extract the signing certificate from a VBA project and write it to a .cer file using Aspose.Cells in C#.
// Use Cases: Verify the authenticity of a macro‑enabled workbook by retrieving its signing certificate. | Create a backup of a VBA project's certificate for compliance or archival purposes. | Automate comparison of certificates across multiple workbooks in a migration or audit scenario.
// AI Prompts: Generate C# code with Aspose.Cells that extracts a VBA project's signing certificate and saves it as a .cer file. | Write a method that checks if a workbook's VBA project is signed and returns the certificate as a byte array. | Provide robust error handling for exporting a VBA certificate when the project is unsigned or the certificate data is missing.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    // Loads a macro‑enabled workbook, accesses its VbaProject, confirms the project is signed, extracts the raw certificate bytes into a MemoryStream, and writes the stream to a .cer file on disk using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Path where the extracted certificate will be saved
            string certificatePath = "VbaCertificate.cer";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the VBA project is signed and certificate data exists
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Export the certificate raw data to a MemoryStream
                using (MemoryStream certStream = new MemoryStream(vbaProject.CertRawData))
                {
                    // Write the MemoryStream contents to a file on disk
                    using (FileStream fileStream = new FileStream(certificatePath, FileMode.Create, FileAccess.Write))
                    {
                        certStream.CopyTo(fileStream);
                    }
                }

                Console.WriteLine($"Certificate exported successfully to '{certificatePath}'.");
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or no certificate data is available.");
            }
        }
    }
}

// Title: Export a VBA project certificate to a .cer file with Aspose.Cells for .NET (C#)
// Description: This example shows how to load a signed .xlsm workbook using Aspose.Cells, access its VbaProject, retrieve the raw certificate bytes (CertRawData) and write them to a .cer file. The code also demonstrates basic error handling and optional workbook saving.
// Keywords: Aspose.Cells | C# | .NET | export VBA certificate | VbaProject CertRawData | signed VBA project | write bytes to file | extract Excel macro certificate | save .cer file | macro project export
// Common Searches: How to export VBA certificate with Aspose.Cells C# | Retrieve CertRawData from a signed .xlsm workbook | Save VBA project certificate as .cer using .NET | Aspose.Cells example for extracting macro signing certificate | Write VBA certificate bytes to disk in C#
// Developer Intent: Extract the signing certificate from a signed VBA project and save it as a .cer file.
// Use Cases: Distribute the VBA signing certificate to external stakeholders for trust verification. | Archive the certificate for compliance or audit purposes after processing a workbook. | Validate certificate size or content before embedding it in another document or system.
// AI Prompts: Generate C# code that loads a signed .xlsm file with Aspose.Cells, checks the VBA project signature, and writes the certificate to a .cer file. | Provide best‑practice error handling for exporting a VBA certificate, including missing file and unsigned project scenarios. | Explain how to programmatically verify the integrity of the exported .cer file after writing it with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // This example shows how to load a signed .xlsm workbook using Aspose.Cells, access its VbaProject, retrieve the raw certificate bytes (CertRawData) and write them to a .cer file. The code also demonstrates basic error handling and optional workbook saving.
    public class ExportVbaCertificate
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string signedWorkbookPath = "SignedWorkbook.xlsm";

                // Verify the source file exists
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"File not found: {signedWorkbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed and certificate data is available
                if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
                {
                    // Retrieve the raw certificate bytes
                    byte[] certificateBytes = vbaProject.CertRawData;

                    // Define the output file name for the exported certificate
                    string outputCertificatePath = "VbaCertificate.cer";

                    // Write the certificate bytes to the file system
                    File.WriteAllBytes(outputCertificatePath, certificateBytes);

                    Console.WriteLine($"Certificate exported successfully to '{outputCertificatePath}'.");
                    Console.WriteLine($"Certificate size: {certificateBytes.Length} bytes.");
                }
                else
                {
                    Console.WriteLine("The VBA project is not signed or no certificate data is available.");
                }

                // Optionally save the workbook after processing
                workbook.Save("ProcessedWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}

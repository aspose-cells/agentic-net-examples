// Title: Export VBA Project Digital Certificate to a MemoryStream (C#) – Aspose.Cells
// Description: Loads a signed .xlsm workbook with Aspose.Cells, accesses its VbaProject, verifies the presence of a digital signature, and creates a MemoryStream from the certificate's raw data for downstream processing such as X509Certificate2 validation.
// Keywords: Aspose.Cells VBA certificate export | C# MemoryStream from VbaProject | extract signed VBA project certificate | CertRawData Aspose.Cells | load X509Certificate2 from Excel VBA | signed .xlsm workbook handling
// Common Searches: how to get VBA project certificate as MemoryStream using Aspose.Cells | C# extract digital signature from Excel macro project | Aspose.Cells read CertRawData from signed workbook | export VBA digital certificate to stream .NET | retrieve VBA project signature with Aspose.Cells
// Developer Intent: Retrieve the digital certificate of a signed VBA project and provide it as a MemoryStream for further cryptographic operations.
// Use Cases: Validate the authenticity of a signed .xlsm file by extracting its VBA certificate. | Convert the MemoryStream containing the certificate into an X509Certificate2 object for thumbprint or expiration checks. | Programmatically determine whether a VBA project is signed before attempting certificate extraction.
// AI Prompts: Generate C# code that opens a signed .xlsm file with Aspose.Cells, checks VbaProject.IsSigned, and returns the certificate as a MemoryStream. | Show how to load the MemoryStream from VbaProject.CertRawData into a System.Security.Cryptography.X509Certificates.X509Certificate2 and verify its thumbprint. | Create a robust method that extracts CertRawData from a VbaProject, handles null or empty data, and outputs a ready‑to‑use MemoryStream.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed .xlsm workbook with Aspose.Cells, accesses its VbaProject, verifies the presence of a digital signature, and creates a MemoryStream from the certificate's raw data for downstream processing such as X509Certificate2 validation.
    public class ExportVbaCertificateToMemoryStream
    {
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"File not found: {signedWorkbookPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project from the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Check if the VBA project is signed and certificate data is available
                if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
                {
                    // Export the certificate raw data to a MemoryStream
                    using (MemoryStream certStream = new MemoryStream(vbaProject.CertRawData))
                    {
                        Console.WriteLine($"Certificate exported to MemoryStream. Length: {certStream.Length}");

                        // Reset position if you need to read from the beginning
                        certStream.Position = 0;

                        // Example of further processing (e.g., loading into X509Certificate2)
                        // var certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(certStream.ToArray());
                    }
                }
                else
                {
                    Console.WriteLine("VBA project is not signed or certificate data is unavailable.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ExportVbaCertificateToMemoryStream.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}

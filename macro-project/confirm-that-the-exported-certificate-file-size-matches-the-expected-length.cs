// Title: C# – Verify Exported VBA Project Certificate Size with Aspose.Cells for .NET
// Description: Loads a signed XLSM workbook, extracts the VBA project's raw certificate data (CertRawData), saves it as a .cer file, and checks that the exported file size matches the original byte array length.
// Keywords: Aspose.Cells VBA certificate export | C# verify certificate file size | CertRawData length comparison | export .cer from signed workbook | .NET VBA project signing validation | Aspose.Cells VbaProject IsSigned | file size integrity check
// Common Searches: How to export a VBA project's certificate with Aspose.Cells and confirm its size | C# compare CertRawData length to exported .cer file size | Aspose.Cells verify signed VBA project before exporting certificate | Check exported certificate file size matches original bytes Aspose.Cells | Export VBA certificate and validate integrity .NET
// Developer Intent: Confirm that the .cer file created from a signed VBA project has exactly the same byte length as the embedded certificate data.
// Use Cases: Automated integrity verification of exported VBA certificates in deployment pipelines. | Compliance testing to ensure no data loss when extracting certificates from signed workbooks. | Diagnostic tool for detecting corrupted or incomplete certificate exports.
// AI Prompts: Generate a C# method that loads a signed workbook, extracts CertRawData, writes it to a .cer file, and returns a boolean indicating size match. | Create robust error handling for missing files, unsigned VBA projects, or empty certificate data when exporting a VBA certificate with Aspose.Cells. | Write an MSTest unit test that asserts the exported certificate file size equals the CertRawData length for a given signed XLSM file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    // Loads a signed XLSM workbook, extracts the VBA project's raw certificate data (CertRawData), saves it as a .cer file, and checks that the exported file size matches the original byte array length.
    public class VerifyExportedCertificateSize
    {
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Ensure the workbook file exists
            if (!File.Exists(signedWorkbookPath))
            {
                Console.WriteLine($"Workbook file not found: {signedWorkbookPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook (uses the load rule)
                workbook = new Workbook(signedWorkbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Ensure the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate to export.");
                return;
            }

            // Get the raw certificate data (byte array)
            byte[] certData = vbaProject.CertRawData;

            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // Export the certificate to a file
            string exportedCertPath = "ExportedCertificate.cer";
            try
            {
                File.WriteAllBytes(exportedCertPath, certData);
                Console.WriteLine($"Certificate exported to: {exportedCertPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write certificate file: {ex.Message}");
                return;
            }

            // Get the file size of the exported certificate
            long fileSize = new FileInfo(exportedCertPath).Length;

            // Compare the file size with the original byte array length
            bool sizeMatches = fileSize == certData.Length;

            Console.WriteLine($"Original certificate byte length: {certData.Length}");
            Console.WriteLine($"Exported file size (bytes): {fileSize}");
            Console.WriteLine($"Size matches expected length: {sizeMatches}");
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                VerifyExportedCertificateSize.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VerifyExportedCertificateSizeDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the workbook that contains a signed VBA project
                string signedWorkbookPath = "SignedWithVba.xlsm";

                // Ensure the file exists before proceeding
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"File not found: {signedWorkbookPath}");
                    return;
                }

                // Load the signed workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is actually signed
                if (!vbaProject.IsSigned)
                {
                    Console.WriteLine("The VBA project is not signed. No certificate to export.");
                    return;
                }

                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Guard against null or empty data
                if (certData == null || certData.Length == 0)
                {
                    Console.WriteLine("Certificate raw data is empty.");
                    return;
                }

                // Define the output certificate file path
                string certFilePath = "ExportedVbaCertificate.cer";

                // Write the certificate bytes to a file
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate exported to: {certFilePath}");

                // Get the file size of the exported certificate
                long fileSize = new FileInfo(certFilePath).Length;

                // Compare the file size with the original byte array length
                bool sizesMatch = fileSize == certData.Length;

                // Output the verification result
                Console.WriteLine($"Original certificate byte length: {certData.Length}");
                Console.WriteLine($"Exported file size (bytes): {fileSize}");
                Console.WriteLine($"Sizes match: {sizesMatch}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VerifyExportedCertificateSizeDemo.Run();
        }
    }
}
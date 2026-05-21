using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateDemo
    {
        public static void Run()
        {
            try
            {
                // Path to the signed workbook containing a VBA project
                string signedWorkbookPath = "SignedWorkbook.xlsm";

                // Ensure the workbook file exists before attempting to load it
                if (!File.Exists(signedWorkbookPath))
                {
                    Console.WriteLine($"Workbook not found: {signedWorkbookPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(signedWorkbookPath);

                // Access the VBA project within the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Verify that the VBA project is signed
                if (!vbaProject.IsSigned)
                {
                    Console.WriteLine("VBA project is not signed. No certificate to export.");
                    return;
                }

                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData == null || certData.Length == 0)
                {
                    Console.WriteLine("Certificate data is empty.");
                    return;
                }

                // Define the output file path for the exported certificate
                string certFilePath = "VbaProjectCertificate.cer";

                // Write the certificate bytes to the file
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate exported to: {certFilePath}");

                // Verify that the file was created successfully
                bool fileExists = File.Exists(certFilePath);
                Console.WriteLine($"Verification - file exists: {fileExists}");
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
            try
            {
                ExportVbaCertificateDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
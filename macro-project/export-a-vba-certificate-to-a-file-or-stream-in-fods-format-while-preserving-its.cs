using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificateToFods
    {
        public static void Run()
        {
            // Path to the source macro-enabled workbook (must be signed to have a certificate)
            string sourcePath = "SignedWorkbook.xlsm";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project (read‑only property)
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Save the certificate to a .cer file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to: {certFilePath}");
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

            // Export the workbook (including its VBA project) to FODS format
            string fodsPath = "SignedWorkbook.fods";
            workbook.Save(fodsPath, SaveFormat.Fods);
            Console.WriteLine($"Workbook saved in FODS format to: {fodsPath}");

            // Also save the workbook to a memory stream in FODS format
            using (MemoryStream fodsStream = new MemoryStream())
            {
                workbook.Save(fodsStream, SaveFormat.Fods);
                fodsStream.Position = 0;
                Console.WriteLine($"Workbook also saved to a memory stream (length: {fodsStream.Length} bytes).");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificateToFods.Run();
        }
    }
}
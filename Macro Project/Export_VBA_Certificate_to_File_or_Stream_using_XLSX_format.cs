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
            // Path to an existing workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Load the workbook (uses the Workbook(string) constructor)
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project (Workbook.VbaProject property)
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data (VbaProject.CertRawData property)
                byte[] certData = vbaProject.CertRawData;

                // If certificate data exists, export it to a file
                if (certData != null && certData.Length > 0)
                {
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");
                }

                // Also demonstrate exporting the certificate to a memory stream
                using (MemoryStream certStream = new MemoryStream())
                {
                    certStream.Write(vbaProject.CertRawData, 0, vbaProject.CertRawData.Length);
                    // Reset position for any further processing
                    certStream.Position = 0;
                    Console.WriteLine($"Certificate written to memory stream (length: {certStream.Length} bytes).");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project.");
            }

            // Save the workbook in XLSX format (Workbook.Save(string, SaveFormat) method)
            string outputPath = "ExportedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved as XLSX: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificateDemo.Run();
        }
    }
}
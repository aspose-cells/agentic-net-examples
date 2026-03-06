using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateToPdfDemo
    {
        public static void Run()
        {
            // Path to a macro-enabled workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the certificate raw data
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists
                if (certData != null && certData.Length > 0)
                {
                    // Save certificate to a physical file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // Also demonstrate saving the certificate to a memory stream
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        // Reset position for potential further processing
                        certStream.Position = 0;
                        Console.WriteLine($"Certificate written to memory stream (length: {certStream.Length} bytes)");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate raw data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
            }

            // Export the workbook (including the VBA project) to PDF format
            string pdfOutputPath = "WorkbookWithVba.pdf";
            workbook.Save(pdfOutputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {pdfOutputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificateToPdfDemo.Run();
        }
    }
}
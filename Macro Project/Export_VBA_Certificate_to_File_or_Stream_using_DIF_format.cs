using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Saving; // For DifSaveOptions

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateDemo
    {
        public static void Run()
        {
            // Path to the workbook that contains a signed VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the certificate raw data
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists
                if (certData != null && certData.Length > 0)
                {
                    // ----- Export certificate to a file -----
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // ----- Export certificate to a memory stream -----
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        // Reset position for potential further reading
                        certStream.Position = 0;
                        Console.WriteLine($"Certificate written to memory stream (length = {certStream.Length} bytes).");
                        // Example: you could return this stream or process it further here
                    }
                }
                else
                {
                    Console.WriteLine("Certificate raw data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project.");
            }

            // ----- Save the workbook in DIF format using DifSaveOptions -----
            DifSaveOptions difOptions = new DifSaveOptions
            {
                ClearData = false,
                CreateDirectory = true,
                RefreshChartCache = true
            };

            string difFilePath = "WorkbookExport.dif";
            workbook.Save(difFilePath, difOptions);
            Console.WriteLine($"Workbook saved in DIF format: {difFilePath}");
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
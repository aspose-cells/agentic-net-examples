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

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the certificate raw data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Save the certificate to a .cer file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // Also write the certificate to a memory stream (example of stream usage)
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        certStream.Position = 0; // Reset position for potential further processing
                        Console.WriteLine($"Certificate written to memory stream (length: {certStream.Length} bytes).");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project.");
            }

            // Save the workbook in StarOffice Calc (SXC) format to a file
            string sxcFilePath = "WorkbookExported.sxc";
            workbook.Save(sxcFilePath, SaveFormat.Sxc);
            Console.WriteLine($"Workbook saved in SXC format to: {sxcFilePath}");

            // Additionally, save the workbook in SXC format to a memory stream
            using (MemoryStream sxcStream = new MemoryStream())
            {
                workbook.Save(sxcStream, SaveFormat.Sxc);
                sxcStream.Position = 0; // Reset for any further use
                Console.WriteLine($"Workbook saved to memory stream in SXC format (length: {sxcStream.Length} bytes).");
            }
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
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    public class ExportVbaCertificate
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
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Export certificate to a physical file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // Export certificate to a memory stream
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        certStream.Position = 0; // Reset for potential further use

                        // Example: write the stream content to another file
                        string streamOutputPath = "VbaCertificateFromStream.cer";
                        using (FileStream fileOut = new FileStream(streamOutputPath, FileMode.Create, FileAccess.Write))
                        {
                            certStream.CopyTo(fileOut);
                        }
                        Console.WriteLine($"Certificate saved from stream to file: {streamOutputPath}");
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

            // Save the workbook as an XLTX template
            string templatePath = "WorkbookTemplate.xltx";
            workbook.Save(templatePath, SaveFormat.Xltx);
            Console.WriteLine($"Workbook saved as XLTX template: {templatePath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportVbaCertificate.Run();
        }
    }
}
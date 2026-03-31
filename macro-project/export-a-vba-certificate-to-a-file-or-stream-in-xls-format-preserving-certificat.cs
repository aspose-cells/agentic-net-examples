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
            // Path to the source workbook that contains a signed VBA project
            string sourcePath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists before proceeding
                if (certData != null && certData.Length > 0)
                {
                    // Save the certificate raw data to a .cer file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to '{certFilePath}' (Length: {certData.Length} bytes).");
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

            // Prepare XLS save options for Excel 97-2003 format
            XlsSaveOptions xlsOptions = new XlsSaveOptions();

            // Save the workbook to a memory stream in XLS format
            using (MemoryStream xlsStream = new MemoryStream())
            {
                workbook.Save(xlsStream, xlsOptions);
                xlsStream.Position = 0; // Reset stream position before copying

                // Write the stream content to a physical .xls file
                string xlsFilePath = "ExportedWorkbook.xls";
                using (FileStream fileStream = new FileStream(xlsFilePath, FileMode.Create, FileAccess.Write))
                {
                    xlsStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Workbook saved as XLS to '{xlsFilePath}'.");
            }
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
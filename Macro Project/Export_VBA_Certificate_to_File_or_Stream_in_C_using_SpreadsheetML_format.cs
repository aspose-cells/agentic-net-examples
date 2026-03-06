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
            // Load a workbook that contains a signed VBA project.
            // Replace the path with the actual file you want to process.
            string signedWorkbookPath = "SignedWorkbook.xlsm";
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed and the certificate data exists.
            if (vbaProject != null && vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Export the certificate raw data to a physical file.
                string certFilePath = "VbaCertificate.cer";
                File.WriteAllBytes(certFilePath, vbaProject.CertRawData);
                Console.WriteLine($"Certificate saved to file: {certFilePath}");

                // Export the same certificate data to a memory stream and then to another file.
                using (MemoryStream certStream = new MemoryStream(vbaProject.CertRawData))
                {
                    string certFromStreamPath = "VbaCertificateFromStream.cer";
                    using (FileStream fileStream = new FileStream(certFromStreamPath, FileMode.Create, FileAccess.Write))
                    {
                        certStream.CopyTo(fileStream);
                    }
                    Console.WriteLine($"Certificate saved from stream to file: {certFromStreamPath}");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or certificate data is unavailable.");
            }

            // Save the workbook in SpreadsheetML (Excel 2003 XML) format.
            string xmlOutputPath = "WorkbookExport.xml";
            workbook.Save(xmlOutputPath, SaveFormat.SpreadsheetML);
            Console.WriteLine($"Workbook saved in SpreadsheetML format: {xmlOutputPath}");
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
using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook that contains a signed VBA project
            string sourcePath = "SignedWorkbook.xlsm";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project (lifecycle rule: property)
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                byte[] certData = vbaProject.CertRawData;

                // ----- Export certificate to a physical file -----
                string certFilePath = "VbaCertificate.cer";
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate raw data saved to file: {certFilePath} (size: {certData.Length} bytes)");

                // ----- Export certificate to a memory stream -----
                using (MemoryStream certStream = new MemoryStream())
                {
                    certStream.Write(certData, 0, certData.Length);
                    // Reset position if further reading is required
                    certStream.Position = 0;
                    Console.WriteLine($"Certificate raw data written to MemoryStream (length: {certStream.Length} bytes)");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
            }

            // ----- Save the workbook as TSV (Tab‑Separated Values) -----
            // Using the Save method with SaveFormat.Tsv (lifecycle rule: save)
            string tsvOutputPath = "WorkbookExport.tsv";
            workbook.Save(tsvOutputPath, SaveFormat.Tsv);
            Console.WriteLine($"Workbook saved in TSV format to: {tsvOutputPath}");
        }
    }
}
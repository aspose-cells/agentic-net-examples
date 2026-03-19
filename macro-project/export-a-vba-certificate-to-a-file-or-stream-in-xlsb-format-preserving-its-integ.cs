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
            // Path to the source workbook that already contains a signed VBA project
            string sourcePath = "SignedWorkbook.xlsm";

            // Load the workbook (create lifecycle)
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project (property lifecycle)
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data (property)
                byte[] certData = vbaProject.CertRawData;

                // Export the certificate to a file (preserving integrity)
                if (certData != null && certData.Length > 0)
                {
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to '{certFilePath}' (Length: {certData.Length} bytes).");
                }

                // Optionally, export the certificate to a memory stream for further processing
                using (MemoryStream certStream = new MemoryStream())
                {
                    certStream.Write(certData, 0, certData.Length);
                    // Reset position if the stream will be read later
                    certStream.Position = 0;
                    // Example: display the first few bytes
                    Console.WriteLine($"First byte of certificate: {certStream.ReadByte()}");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
            }

            // Save the workbook as XLSB while preserving the VBA signature
            XlsbSaveOptions xlsbOptions = new XlsbSaveOptions
            {
                // Example option; adjust as needed
                ExportAllColumnIndexes = true
            };

            string outputPath = "SignedWorkbookExported.xlsb";
            workbook.Save(outputPath, xlsbOptions);
            Console.WriteLine($"Workbook saved as XLSB to '{outputPath}'.");
        }
    }
}
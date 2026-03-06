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
            // Path to an existing macro-enabled workbook that is signed.
            // Replace with the actual path of your signed .xlsm file.
            string signedWorkbookPath = "SignedWorkbook.xlsm";

            // Load the signed workbook.
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project.
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed.
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data.
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists.
                if (certData != null && certData.Length > 0)
                {
                    // ----- Save certificate to a physical file -----
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate saved to file: {certFilePath}");

                    // ----- Save certificate to a memory stream (optional) -----
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        // Reset position if further processing is needed.
                        certStream.Position = 0;
                        Console.WriteLine($"Certificate written to memory stream (length: {certStream.Length} bytes).");
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

            // ----- Save the workbook in FODS (Flat OpenDocument Spreadsheet) format -----
            // This demonstrates using the Save method with SaveFormat.Fods as required.
            string fodsOutputPath = "WorkbookExported.fods";
            workbook.Save(fodsOutputPath, SaveFormat.Fods);
            Console.WriteLine($"Workbook saved in FODS format: {fodsOutputPath}");
        }
    }
}
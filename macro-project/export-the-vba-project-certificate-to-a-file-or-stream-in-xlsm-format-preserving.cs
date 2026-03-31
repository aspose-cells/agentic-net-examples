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
            // Path to the macro‑enabled workbook that contains a signed VBA project
            string inputPath = "SignedWorkbook.xlsm";

            // Load the workbook (create rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project (property rule)
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data (property rule)
                byte[] certData = vbaProject.CertRawData;

                // Ensure certificate data exists
                if (certData != null && certData.Length > 0)
                {
                    // Export the certificate to a .cer file
                    string certFilePath = "VbaProjectCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate exported to: {certFilePath} (size: {certData.Length} bytes)");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not signed; no certificate to export.");
            }

            // Optionally, save the workbook back as XLSM (save rule)
            string outputPath = "SignedWorkbook_Exported.xlsm";
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
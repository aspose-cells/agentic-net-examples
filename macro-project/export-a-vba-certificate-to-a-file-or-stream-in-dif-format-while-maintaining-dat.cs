using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Saving;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string sourceWorkbookPath = "SignedWorkbook.xlsm";      // Workbook that contains a signed VBA project
            string certificateOutputPath = "VbaCertificate.cer";   // File to store the extracted certificate
            string difOutputPath = "WorkbookExport.dif";           // DIF file output path

            ExportVbaCertificateAndSaveDif(sourceWorkbookPath, certificateOutputPath, difOutputPath);
        }

        /// <summary>
        /// Loads a workbook, extracts the VBA project's certificate (if signed),
        /// writes the certificate raw data to a file, and saves the workbook in DIF format.
        /// </summary>
        static void ExportVbaCertificateAndSaveDif(string workbookPath, string certPath, string difPath)
        {
            // Load the workbook (macro‑enabled format is required for VBA)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed before attempting to read the certificate
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate bytes
                byte[] certData = vbaProject.CertRawData;

                // Ensure we have data before writing
                if (certData != null && certData.Length > 0)
                {
                    // Write the certificate raw data to the specified file
                    File.WriteAllBytes(certPath, certData);
                    Console.WriteLine($"Certificate extracted and saved to '{certPath}'. Length: {certData.Length} bytes.");
                }
                else
                {
                    Console.WriteLine("Certificate data is empty; nothing was written.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project; certificate extraction skipped.");
            }

            // Prepare DIF save options – customize as required
            DifSaveOptions difOptions = new DifSaveOptions
            {
                // Example settings; adjust based on your needs
                ClearData = false,
                CreateDirectory = true,
                RefreshChartCache = true
            };

            // Save the workbook in DIF format using the provided Save method with options
            workbook.Save(difPath, difOptions);
            Console.WriteLine($"Workbook saved in DIF format to '{difPath}'.");
        }
    }
}
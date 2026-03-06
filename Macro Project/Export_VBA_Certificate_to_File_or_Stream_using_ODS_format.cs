using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Ods;

namespace AsposeCellsVbaCertificateExport
{
    public class Exporter
    {
        public static void ExportVbaCertificate(string workbookPath)
        {
            // Load the workbook that contains a signed VBA project
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed and the certificate data is available
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                byte[] certData = vbaProject.CertRawData;

                // Save the certificate raw data to a .cer file
                string certFilePath = Path.ChangeExtension(workbookPath, ".cer");
                File.WriteAllBytes(certFilePath, certData);
                Console.WriteLine($"Certificate saved to file: {certFilePath}");

                // Also write the certificate data to a memory stream (for further processing if needed)
                using (MemoryStream certStream = new MemoryStream(certData))
                {
                    // Example: display the length of the stream
                    Console.WriteLine($"Certificate stream length: {certStream.Length} bytes");
                    // The stream can be used elsewhere as required
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
            }

            // Prepare ODS save options (e.g., specify LibreOffice generator and ODF version 1.2)
            OdsSaveOptions odsOptions = new OdsSaveOptions();
            odsOptions.GeneratorType = OdsGeneratorType.LibreOffice;
            odsOptions.OdfStrictVersion = OpenDocumentFormatVersionType.Odf12;

            // Save the workbook as an ODS file on disk
            string odsFilePath = Path.ChangeExtension(workbookPath, ".ods");
            workbook.Save(odsFilePath, odsOptions);
            Console.WriteLine($"Workbook saved as ODS: {odsFilePath}");

            // Additionally, save the ODS content to a memory stream
            using (MemoryStream odsStream = new MemoryStream())
            {
                workbook.Save(odsStream, odsOptions);
                // Reset position if the stream will be read later
                odsStream.Position = 0;
                Console.WriteLine($"ODS content saved to memory stream. Length: {odsStream.Length} bytes");
                // The stream can be returned, sent over network, etc.
            }
        }

        // Example usage
        public static void Main()
        {
            // Path to the source workbook (must be macro-enabled and signed)
            string sourceWorkbook = "SignedWorkbook.xlsm";

            if (!File.Exists(sourceWorkbook))
            {
                Console.WriteLine($"Source workbook not found: {sourceWorkbook}");
                return;
            }

            ExportVbaCertificate(sourceWorkbook);
        }
    }
}
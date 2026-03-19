using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    public class ExportVbaCertificateToPdf
    {
        public static void Run()
        {
            // Path to an existing workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed and certificate data is available
            if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                // Retrieve the raw certificate bytes
                byte[] certData = vbaProject.CertRawData;

                // Create a temporary certificate file (required for OLE embedding)
                string tempCertPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".cer");
                File.WriteAllBytes(tempCertPath, certData);

                // Read the file bytes for OLE insertion
                byte[] oleData = File.ReadAllBytes(tempCertPath);

                // Insert the certificate as an OLE object into the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                // Add the OLE object (row, column, height, width, ole data)
                int oleIndex = sheet.OleObjects.Add(5, 0, 200, 50, oleData);
                // Display the OLE object as an icon for better PDF appearance
                sheet.OleObjects[oleIndex].DisplayAsIcon = true;

                // Clean up the temporary file
                File.Delete(tempCertPath);

                // Prepare PDF save options with attachment embedding enabled
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    EmbedAttachments = true,
                    ExportDocumentStructure = true,
                    CustomPropertiesExport = PdfCustomPropertiesExport.Standard
                };

                // Save the workbook as PDF
                string outputPdfPath = "VbaCertificateExport.pdf";
                workbook.Save(outputPdfPath, pdfOptions);

                Console.WriteLine($"Certificate exported and embedded in PDF: {outputPdfPath}");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a signed VBA project or certificate data is unavailable.");
            }

            // Dispose resources
            workbook.Dispose();
        }
    }

    public class Program
    {
        public static void Main()
        {
            ExportVbaCertificateToPdf.Run();
        }
    }
}
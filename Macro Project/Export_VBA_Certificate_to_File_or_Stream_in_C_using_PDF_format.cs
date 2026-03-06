using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main(string[] args)
        {
            ExportCertificateToPdf.Run();
        }
    }

    public class ExportCertificateToPdf
    {
        public static void Run()
        {
            // Path to the source Excel file that contains a signed VBA project
            string sourceExcelPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(sourceExcelPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project is signed
            if (!vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is not signed. No certificate to export.");
                return;
            }

            // Retrieve the raw certificate data
            byte[] certData = vbaProject.CertRawData;
            if (certData == null || certData.Length == 0)
            {
                Console.WriteLine("Certificate raw data is empty.");
                return;
            }

            // Convert the binary certificate data to a Base64 string for readable representation
            string certBase64 = Convert.ToBase64String(certData);

            // Create a new worksheet to hold the certificate information
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet certSheet = workbook.Worksheets[sheetIndex];
            certSheet.Name = "VBA Certificate";

            // Write a header
            certSheet.Cells["A1"].PutValue("VBA Project Certificate (Base64)");

            // Write the certificate data (split into multiple rows for readability)
            const int charsPerRow = 80;
            for (int i = 0, row = 2; i < certBase64.Length; i += charsPerRow, row++)
            {
                string segment = certBase64.Substring(i, Math.Min(charsPerRow, certBase64.Length - i));
                certSheet.Cells[row, 0].PutValue(segment);
            }

            // Export to PDF file
            string pdfFilePath = "VbaCertificate.pdf";
            workbook.Save(pdfFilePath, SaveFormat.Pdf);
            Console.WriteLine($"Certificate exported to PDF file: {pdfFilePath}");

            // Export to PDF stream
            using (MemoryStream pdfStream = new MemoryStream())
            {
                workbook.Save(pdfStream, SaveFormat.Pdf);
                File.WriteAllBytes("VbaCertificate_Stream.pdf", pdfStream.ToArray());
                Console.WriteLine("Certificate exported to PDF stream and saved as VbaCertificate_Stream.pdf");
            }
        }
    }
}
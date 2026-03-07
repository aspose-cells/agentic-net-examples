using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfA1aConversion
{
    public class Converter
    {
        public static void ConvertExcelToPdfA1a(string sourcePath, string destinationPath)
        {
            // Load the existing Excel workbook
            Workbook workbook = new Workbook(sourcePath);

            // Create PDF save options and set compliance to PDF/A-1a
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the workbook as a PDF/A-1a compliant file
            workbook.Save(destinationPath, pdfOptions);
        }

        // Example usage
        public static void Main()
        {
            string excelFile = "input.xlsx";
            string pdfFile = "output.pdfa1a.pdf";

            ConvertExcelToPdfA1a(excelFile, pdfFile);

            Console.WriteLine("Conversion to PDF/A-1a completed successfully.");
        }
    }
}
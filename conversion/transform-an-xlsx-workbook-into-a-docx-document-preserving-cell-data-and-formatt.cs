using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class XlsxToDocxConverter
    {
        public static void Convert(string sourceXlsxPath, string destinationDocxPath)
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(sourceXlsxPath);

            // Initialize DOCX save options (preserves formatting by default)
            DocxSaveOptions docxOptions = new DocxSaveOptions();

            // Save the workbook as a DOCX document
            workbook.Save(destinationDocxPath, docxOptions);

            Console.WriteLine($"Conversion completed: '{sourceXlsxPath}' → '{destinationDocxPath}'");
        }

        // Example usage
        public static void Main()
        {
            string sourceFile = "sample.xlsx";
            string destFile = "sample.docx";

            Convert(sourceFile, destFile);
        }
    }
}
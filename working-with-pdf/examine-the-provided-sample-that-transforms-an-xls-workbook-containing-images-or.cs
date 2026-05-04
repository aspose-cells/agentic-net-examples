using System;
using Aspose.Cells;

namespace AsposeCellsPdfConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelToPdfConverter.Run();
        }
    }

    public class ExcelToPdfConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xls";
            string outputPath = "output.pdf";

            Workbook workbook = new Workbook(sourcePath);
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                EmbedAttachments = true
            };
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Excel file '{sourcePath}' has been successfully converted to PDF at '{outputPath}'.");
        }
    }
}
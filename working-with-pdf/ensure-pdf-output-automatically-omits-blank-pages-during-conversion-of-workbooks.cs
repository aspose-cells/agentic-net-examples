using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class OmitBlankPagesPdfDemo
    {
        public static void Run()
        {
            // Create a new workbook (default workbook contains one worksheet)
            Workbook workbook = new Workbook();

            // Add minimal printable content to avoid "nothing to print" exception
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(" "); // a space ensures printable content

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false,
                PrintingPageType = PrintingPageType.IgnoreBlank
            };

            // Define output file path
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Workbook_NoBlankPages.pdf");

            // Save the workbook to PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved without blank pages: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OmitBlankPagesPdfDemo.Run();
        }
    }
}
using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

namespace AsposeCellsPdfPerSheet
{
    public class WorkbookToPdfPerWorksheet
    {
        public static void Run()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options so each worksheet is rendered on a separate page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true   // Each sheet becomes one PDF page
            };

            // Save the workbook as a PDF file (replace with your desired output path)
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook has been saved to PDF with each worksheet on its own page: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToPdfPerWorksheet.Run();
        }
    }
}
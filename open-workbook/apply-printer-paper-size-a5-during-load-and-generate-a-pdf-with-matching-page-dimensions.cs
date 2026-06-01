using System;
using Aspose.Cells;

namespace AsposeCellsA5PdfDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (replace with your actual file path)
            string sourceFile = "input.xlsx";

            // Create LoadOptions and set the default printer paper size to A5
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.SetPaperSize(PaperSizeType.PaperA5);

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Ensure each worksheet uses A5 paper size (optional, but guarantees the PDF matches)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
            }

            // Save the workbook as PDF; the page dimensions will correspond to A5 size
            string outputPdf = "output_A5.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved as PDF with A5 page size: {outputPdf}");
        }
    }
}
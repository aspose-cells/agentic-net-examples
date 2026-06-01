using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace MergeWorksheetsToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file containing multiple worksheets
            string excelPath = "input.xlsx";

            // Desired output PDF file path
            string pdfPath = "merged_output.pdf";

            // Load the workbook (creates a Workbook instance and loads the file)
            Workbook workbook = new Workbook(excelPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Preserve the original order of all worksheets by selecting the full sheet set
            // SheetSet.All returns a set with all sheets in their original order
            pdfOptions.SheetSet = SheetSet.All;

            // Save the workbook as a single PDF containing all worksheets
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Workbook sheets merged into PDF successfully: {pdfPath}");
        }
    }
}
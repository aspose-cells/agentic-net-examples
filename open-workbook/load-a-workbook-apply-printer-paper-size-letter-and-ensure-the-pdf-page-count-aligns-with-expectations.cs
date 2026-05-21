using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPaperSizeDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Apply the default printer paper size (Letter) to the workbook settings
            workbook.Settings.PaperSize = PaperSizeType.PaperLetter;

            // Also set the paper size for each worksheet's PageSetup (optional but ensures consistency)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
            }

            // Create print options – we will use them for preview and PDF export
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                // Ensure the output format is PDF
                SaveFormat = SaveFormat.Pdf
            };

            // Evaluate the total page count after applying the paper size settings
            WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, printOptions);
            int expectedPageCount = preview.EvaluatedPageCount;
            Console.WriteLine($"Evaluated PDF page count: {expectedPageCount}");

            // Save the workbook as PDF
            string outputPdfPath = "output.pdf";
            workbook.Save(outputPdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to: {outputPdfPath}");

            // (Optional) Verify the page count using WorkbookRender after saving
            WorkbookRender render = new WorkbookRender(workbook, printOptions);
            int actualPageCount = render.PageCount;
            Console.WriteLine($"Actual PDF page count reported by WorkbookRender: {actualPageCount}");

            // Clean up resources
            render.Dispose();
        }
    }
}
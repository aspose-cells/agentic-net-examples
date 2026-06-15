using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsOutlinePdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "OutlineDemo";

            // Populate some sample data
            for (int row = 0; row < 10; row++)
            {
                sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                sheet.Cells[row, 1].PutValue((row + 1) * 10);
            }

            // Create first level row group (rows 1‑4)
            sheet.Cells.GroupRows(0, 3, false);
            // Create nested row group inside the first group (rows 2‑3)
            sheet.Cells.GroupRows(1, 2, true);

            // Create first level column group (columns A‑B)
            sheet.Cells.GroupColumns(0, 1, false);
            // Create nested column group inside the first group (column B only)
            sheet.Cells.GroupColumns(1, 1, true);

            // Ensure the outline is visible in the worksheet
            sheet.IsOutlineShown = true;

            // Configure PDF save options to export document structure (outline)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF; the resulting PDF will contain an outline
            // that mirrors the nested row and column groups defined above.
            string outputPath = "OutlineExported.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to '{outputPath}' with document structure exported.");
        }
    }
}
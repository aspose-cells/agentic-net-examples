using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsGroupColumnsAutoFitRowsPdf
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data across several columns and rows
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1} - Some long text to demonstrate autofit");
                }
            }

            // Group columns C (index 2) to F (index 5)
            cells.GroupColumns(2, 5);

            // Auto‑fit all rows so that row heights adjust to the content within the grouped columns
            worksheet.AutoFitRows();

            // Optional: ensure the printed sheet fits all columns on a single page
            worksheet.PageSetup.FitToPagesWide = 1;   // fit columns to one page wide
            worksheet.PageSetup.FitToPagesTall = 0;   // let height adjust automatically

            // Prepare PDF save options (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                OnePagePerSheet = true
            };

            // Export the worksheet to PDF
            workbook.Save("GroupedColumns_AutoFitRows.pdf", pdfOptions);
        }
    }
}
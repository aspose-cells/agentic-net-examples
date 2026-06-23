using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsFreezePanePdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze panes at cell C3 (row index 2, column index 2) with 2 rows and 2 columns frozen
            worksheet.FreezePanes("C3", 2, 2);

            // Prepare PDF save options (default options retain the frozen view)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF while preserving the frozen panes view
            workbook.Save("FrozenPaneOutput.pdf", pdfOptions);
        }
    }
}
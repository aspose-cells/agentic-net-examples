using System;
using Aspose.Cells;

namespace AsposeCellsPdfOutlineDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            for (int i = 0; i < 12; i++)
            {
                sheet.Cells[$"A{i + 1}"].PutValue($"Item {i + 1}");
                sheet.Cells[$"B{i + 1}"].PutValue((i + 1) * 10);
            }

            // Create outline groups:
            // First level group rows 1-6 (0‑based indices 0‑5)
            sheet.Cells.GroupRows(0, 5, false);
            // Second level nested group rows 3-5 (indices 2‑4) inside the first group
            sheet.Cells.GroupRows(2, 4, true);
            // Third level group rows 9-12 (indices 8‑11) as a separate outline
            sheet.Cells.GroupRows(8, 11, false);

            // Ensure the outline is visible in the worksheet view
            sheet.IsOutlineShown = true;

            // Configure PDF save options to export the document structure (outline)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true   // Enables PDF bookmarks reflecting the worksheet outline
            };

            // Save the workbook as PDF with the specified options
            workbook.Save("WorkbookWithOutline.pdf", pdfOptions);

            Console.WriteLine("PDF saved with exported document structure (outline).");
        }
    }
}
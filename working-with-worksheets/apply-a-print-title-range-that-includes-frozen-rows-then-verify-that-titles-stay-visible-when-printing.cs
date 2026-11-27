// Title: Configure print title rows to include frozen panes and verify they repeat on each PDF page with Aspose.Cells for .NET (C#)
// Common Searches: C# example for applying print titles to frozen rows in Aspose.Cells | how to keep header rows visible on every PDF page when using Aspose.Cells | Aspose.Cells .NET code to freeze first rows and repeat them in printed output | checking that frozen panes are part of the print title range before PDF export | setting up page layout with frozen header rows for PDF printing in Aspose.Cells
// Tags: configure print title rows Aspose.Cells C# | freeze panes with print titles Aspose.Cells | export worksheet to PDF preserving print titles Aspose.Cells | validate print title range includes frozen rows C# | page setup print titles PDF Aspose.Cells

using System;
using Aspose.Cells;

namespace PrintTitleWithFrozenRowsDemo
{
    // The sample creates a workbook, populates 100 rows, freezes the first two rows, sets the PrintTitleRows to "1:2", validates that the title range covers the frozen rows, and saves the workbook as a PDF using PdfSaveOptions, ensuring the frozen header rows appear at the top of every printed page.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // 2. Populate the worksheet with sample data (enough rows to span multiple pages)
                const int totalRows = 100;
                for (int row = 0; row < totalRows; row++)
                {
                    if (row == 0) // Header row
                    {
                        sheet.Cells[row, 0].PutValue("Item");
                        sheet.Cells[row, 1].PutValue("Quantity");
                        sheet.Cells[row, 2].PutValue("Price");
                        sheet.Cells[row, 3].PutValue("Total");
                    }
                    else if (row == 1) // Separator row
                    {
                        sheet.Cells[row, 0].PutValue("----");
                        sheet.Cells[row, 1].PutValue("----");
                        sheet.Cells[row, 2].PutValue("----");
                        sheet.Cells[row, 3].PutValue("----");
                    }
                    else // Data rows
                    {
                        sheet.Cells[row, 0].PutValue($"Product {row - 1}");
                        sheet.Cells[row, 1].PutValue(row * 2);
                        sheet.Cells[row, 2].PutValue(5.5);
                        sheet.Cells[row, 3].Formula = $"=B{row + 1}*C{row + 1}";
                    }
                }

                // 3. Freeze the first two rows (rows 1‑2) so they stay visible while scrolling
                // FreezePanes(row, column, freezedRows, freezedColumns)
                // Row/Column specify the cell just below/right of the frozen pane.
                // To freeze rows 1‑2, use row = 2, column = 0, freezedRows = 2, freezedColumns = 0.
                sheet.FreezePanes(2, 0, 2, 0);

                // 4. Apply print title rows that include the frozen rows
                sheet.PageSetup.PrintTitleRows = "1:2";

                // 5. Verify that the title range indeed covers the frozen rows
                string titleRange = sheet.PageSetup.PrintTitleRows;
                if (string.IsNullOrEmpty(titleRange))
                    throw new InvalidOperationException("PrintTitleRows was not set.");

                string[] parts = titleRange.Split(':');
                if (parts.Length != 2 ||
                    !int.TryParse(parts[0], out int startRow) ||
                    !int.TryParse(parts[1], out int endRow))
                    throw new InvalidOperationException($"Unable to parse PrintTitleRows value \"{titleRange}\".");

                int frozenRowCount = 2;
                if (startRow > frozenRowCount)
                    throw new InvalidOperationException("Print title rows do not include the frozen rows.");

                Console.WriteLine($"Print title rows set to \"{titleRange}\" and include the first {frozenRowCount} frozen rows.");

                // 6. Save the workbook to PDF to emulate printing.
                string outputPath = "PrintTitleWithFrozenRows.pdf";
                PdfSaveOptions pdfOptions = new PdfSaveOptions(); // default options respect print titles
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook saved as PDF to \"{outputPath}\".");
                Console.WriteLine("Open the PDF and verify that rows 1‑2 appear at the top of every printed page.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

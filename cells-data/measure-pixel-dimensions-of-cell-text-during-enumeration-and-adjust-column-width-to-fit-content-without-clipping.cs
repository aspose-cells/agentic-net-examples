using System;
using Aspose.Cells;

namespace AsposeCellsColumnAutoFit
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data – in real scenarios the workbook may already contain data
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("This is a considerably longer piece of text");
            cells["B1"].PutValue("Another column");
            cells["B2"].PutValue("Medium length");
            cells["C1"].PutValue("1234567890");
            cells["C2"].PutValue("Short");

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Array to hold the maximum pixel width required for each column
            int[] maxPixelWidth = new int[maxCol + 1];

            // Enumerate each cell, measure its text width in pixels and update column max
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Value != null)
                    {
                        // Get the pixel width of the cell's displayed value
                        int pixelWidth = cell.GetWidthOfValue();

                        // Update the maximum width for this column if needed
                        if (pixelWidth > maxPixelWidth[col])
                            maxPixelWidth[col] = pixelWidth;
                    }
                }
            }

            // Apply the calculated widths to the columns (add a small padding)
            const int padding = 5; // extra pixels to avoid clipping
            for (int col = 0; col <= maxCol; col++)
            {
                // SetColumnWidthPixel sets the column width in normal view (pixels)
                cells.SetColumnWidthPixel(col, maxPixelWidth[col] + padding);
            }

            // Save the workbook
            workbook.Save("ColumnAutoFitByPixel.xlsx");
        }
    }
}
using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthAdjustment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data with varying lengths
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("A much longer piece of text that should expand the column");
            cells["B1"].PutValue(12345);
            cells["B2"].PutValue("Medium length");
            cells["C1"].PutValue("Tiny");
            cells["C2"].PutValue("Another very long text entry to test column width calculation");

            // Determine the range to evaluate (all used rows and columns)
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Array to hold the maximum pixel width per column
            int[] maxPixelWidthPerColumn = new int[maxColumn + 1];

            // Enumerate each cell, measure its value width in pixels, and track the maximum per column
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        // Get the pixel width of the cell's displayed value
                        int pixelWidth = cell.GetWidthOfValue();

                        // Update the maximum width for this column if necessary
                        if (pixelWidth > maxPixelWidthPerColumn[col])
                        {
                            maxPixelWidthPerColumn[col] = pixelWidth;
                        }
                    }
                }
            }

            // Apply the measured widths to the columns (add a small padding for visual comfort)
            const int paddingPixels = 5;
            for (int col = 0; col <= maxColumn; col++)
            {
                int finalWidth = maxPixelWidthPerColumn[col] + paddingPixels;
                // Set column width in normal view using pixel units
                cells.SetColumnWidthPixel(col, finalWidth);
            }

            // Save the workbook
            workbook.Save("AdjustedColumnWidths.xlsx");
        }
    }
}
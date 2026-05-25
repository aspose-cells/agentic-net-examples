using System;
using Aspose.Cells;

class AdjustColumnWidthDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Short");
        cells["A2"].PutValue("This is a longer text that should expand column A");
        cells["B1"].PutValue("Another column with text");
        cells["B2"].PutValue("Short");

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // Store the maximum pixel width required for each column
        int[] maxPixelWidth = new int[maxCol + 1];

        // Enumerate cells, measure pixel width of each value, and track the maximum per column
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type != CellValueType.IsNull)
                {
                    // Get the width of the cell's value in pixels
                    int width = cell.GetWidthOfValue();

                    // Add a small padding to avoid clipping
                    width += 5;

                    if (width > maxPixelWidth[col])
                        maxPixelWidth[col] = width;
                }
            }
        }

        // Apply the calculated column widths
        for (int col = 0; col <= maxCol; col++)
        {
            if (maxPixelWidth[col] > 0)
            {
                // Set column width in pixels
                cells.SetColumnWidthPixel(col, maxPixelWidth[col]);
            }
        }

        // Save the workbook
        workbook.Save("AdjustedColumns.xlsx");
    }
}
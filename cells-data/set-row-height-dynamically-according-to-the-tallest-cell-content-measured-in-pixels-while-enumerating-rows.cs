using System;
using Aspose.Cells;

namespace DynamicRowHeightDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Sample data – in a real scenario the worksheet would already
            // contain data. Here we add some varied text to demonstrate.
            // ------------------------------------------------------------
            cells["A1"].PutValue("Short");
            cells["B1"].PutValue("A longer piece of text that should wrap and increase row height.");
            cells["C1"].PutValue("Medium length");

            cells["A2"].PutValue("Another row with a very very long text that will definitely need more height than the default.");
            cells["B2"].PutValue("Short");

            // Enable text wrapping so that height depends on content length
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;
            cells["A1"].SetStyle(wrapStyle);
            cells["B1"].SetStyle(wrapStyle);
            cells["C1"].SetStyle(wrapStyle);
            cells["A2"].SetStyle(wrapStyle);
            cells["B2"].SetStyle(wrapStyle);

            // Determine the range that actually contains data
            int maxRow = cells.MaxDataRow;      // zero‑based index of last used row
            int maxCol = cells.MaxDataColumn;   // zero‑based index of last used column

            // Iterate through each row and calculate the tallest cell height (in pixels)
            for (int row = 0; row <= maxRow; row++)
            {
                int tallestPixelHeight = 0;

                // Scan all columns in the current row
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Skip empty cells – GetHeightOfValue returns 0 for them
                    if (cell == null || cell.Type == CellValueType.IsNull)
                        continue;

                    // Measure the height required for the cell's value (pixels)
                    int cellHeight = cell.GetHeightOfValue();

                    // Keep the maximum height found in this row
                    if (cellHeight > tallestPixelHeight)
                        tallestPixelHeight = cellHeight;
                }

                // If we found at least one non‑empty cell, set the row height
                if (tallestPixelHeight > 0)
                {
                    // SetRowHeightPixel expects height in pixels
                    cells.SetRowHeightPixel(row, tallestPixelHeight);
                }
            }

            // Save the workbook (adjust the path as needed)
            workbook.Save("DynamicRowHeightOutput.xlsx");
        }
    }
}
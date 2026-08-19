// Title: C# – Auto‑fit Excel column widths by measuring cell text pixels with Aspose.Cells
// Description: Creates a workbook, fills cells with varied text, loops through each column, uses Cell.GetWidthOfValue() to obtain the pixel width of non‑empty cells, records the widest value, adds a small padding, and applies the calculated width with Cells.SetColumnWidthPixel() before saving the file.
// Keywords: Aspose.Cells | C# | SetColumnWidthPixel | GetWidthOfValue | auto fit column width | measure cell text pixel | Excel column resizing | pixel based column width | adjust column width programmatically
// Common Searches: Aspose.Cells auto fit column width C# | GetWidthOfValue example | SetColumnWidthPixel usage | prevent text clipping in Excel with Aspose | measure cell text width Aspose.Cells .NET
// Developer Intent: Automatically resize every worksheet column so the longest cell content fits completely without clipping.
// Use Cases: Generating Excel reports where column widths adapt to dynamic text lengths for optimal readability. | Building templates that guarantee all inserted data remains fully visible after export.
// AI Prompts: Show how to factor font size and style into the width calculation when using GetWidthOfValue. | Provide code that caps column width at a maximum value while still using pixel‑based measurement. | Explain how to apply the same pixel‑measurement technique to auto‑fit row heights with GetHeightOfValue.

using System;
using Aspose.Cells;

// Creates a workbook, fills cells with varied text, loops through each column, uses Cell.GetWidthOfValue() to obtain the pixel width of non‑empty cells, records the widest value, adds a small padding, and applies the calculated width with Cells.SetColumnWidthPixel() before saving the file.
class AdjustColumnWidth
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data to demonstrate column width adjustment
        cells["A1"].PutValue("Short");
        cells["A2"].PutValue("This is a much longer text that should expand the column width");
        cells["B1"].PutValue("Another column");
        cells["B2"].PutValue("Medium length text");
        cells["C1"].PutValue("A very very long piece of text that will require a wide column to avoid clipping");

        // Determine the range of used rows and columns
        int maxRow = cells.MaxDataRow;      // zero‑based index of the last used row
        int maxColumn = cells.MaxDataColumn; // zero‑based index of the last used column

        // Iterate through each column, measure the pixel width of each cell's value,
        // and set the column width to the maximum measured width plus a small padding.
        for (int col = 0; col <= maxColumn; col++)
        {
            int maxPixelWidth = 0;

            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = cells[row, col];

                // Skip empty cells
                if (cell.Value == null)
                    continue;

                // Get the pixel width of the cell's displayed value
                int pixelWidth = cell.GetWidthOfValue();

                // Keep track of the widest cell in the column
                if (pixelWidth > maxPixelWidth)
                    maxPixelWidth = pixelWidth;
            }

            // Add a small padding to avoid clipping due to rounding
            int paddingPixels = 5;
            int finalWidth = maxPixelWidth + paddingPixels;

            // Set the column width in pixels
            cells.SetColumnWidthPixel(col, finalWidth);
        }

        // Save the workbook with adjusted column widths
        workbook.Save("AdjustedColumns.xlsx");
    }
}

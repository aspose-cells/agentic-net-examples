using System;
using Aspose.Cells;

namespace AdjustColumnWidthDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with varying text lengths
            string[] sampleData = {
                "Short",
                "A longer piece of text",
                "Very very long text that needs more column width to be fully visible"
            };

            // Populate column A (index 0) with the sample data
            for (int i = 0; i < sampleData.Length; i++)
            {
                cells[i, 0].PutValue(sampleData[i]);
            }

            // Determine the range that contains data
            int maxColumn = cells.MaxColumn; // last column index that has data
            int maxRow = cells.MaxRow;       // last row index that has data

            // Iterate through each column in the used range
            for (int col = 0; col <= maxColumn; col++)
            {
                int maxPixelWidth = 0;

                // Find the maximum pixel width of cell values in this column
                for (int row = 0; row <= maxRow; row++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        int pixelWidth = cell.GetWidthOfValue(); // width of the cell's value in pixels
                        if (pixelWidth > maxPixelWidth)
                        {
                            maxPixelWidth = pixelWidth;
                        }
                    }
                }

                // If the column contains any data, set its width based on the measured pixel width
                if (maxPixelWidth > 0)
                {
                    // Add a small padding (e.g., 5 pixels) to avoid clipping
                    int paddedWidth = maxPixelWidth + 5;
                    cells.SetColumnWidthPixel(col, paddedWidth);
                }
            }

            // Save the workbook with adjusted column widths
            workbook.Save("AdjustedColumnWidth.xlsx");
        }
    }
}
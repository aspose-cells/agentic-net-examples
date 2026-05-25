using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightAdjustment
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with long wrapped text
            cells["A1"].PutValue("This is a very long text that should wrap inside the cell and cause the row height to increase accordingly.");
            cells["B1"].PutValue("Short text");
            cells["A2"].PutValue("Another long piece of text that will be wrapped and needs proper row height.");
            cells["B2"].PutValue("Another short");

            // Enable text wrapping for the cells that contain long text
            Style wrapStyle = worksheet.Cells["A1"].GetStyle();
            wrapStyle.IsTextWrapped = true;
            worksheet.Cells["A1"].SetStyle(wrapStyle);
            worksheet.Cells["A2"].SetStyle(wrapStyle);

            // Determine the maximum row index that contains data
            int maxRow = cells.MaxDataRow;

            // Iterate through each row to calculate required pixel height
            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                int maxPixelHeight = 0;

                // Iterate through each column in the current row
                for (int colIndex = 0; colIndex <= cells.MaxDataColumn; colIndex++)
                {
                    Cell cell = cells[rowIndex, colIndex];
                    if (cell == null) continue;

                    // Check if the cell has text wrapping enabled
                    Style cellStyle = cell.GetStyle();
                    if (cellStyle.IsTextWrapped)
                    {
                        // Get the height required for the cell's value in pixels
                        int cellPixelHeight = cell.GetHeightOfValue();

                        // Keep the maximum height among cells in the row
                        if (cellPixelHeight > maxPixelHeight)
                            maxPixelHeight = cellPixelHeight;
                    }
                }

                // If any wrapped cell was found, set the row height in pixels
                if (maxPixelHeight > 0)
                {
                    // SetRowHeightPixel expects height in pixels
                    cells.SetRowHeightPixel(rowIndex, maxPixelHeight);
                }
            }

            // Save the workbook to a file
            workbook.Save("RowHeightAdjusted.xlsx");
        }
    }
}
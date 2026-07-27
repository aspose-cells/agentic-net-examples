using System;
using Aspose.Cells;

namespace AsposeCellsDynamicRowHeight
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data with varying text lengths to demonstrate dynamic row height
            cells["A1"].PutValue("Short");
            cells["B1"].PutValue("This is a longer piece of text that should increase the row height.");
            cells["A2"].PutValue("Another short");
            cells["B2"].PutValue("Multi-line\ntext\nexample");
            cells["C3"].PutValue("A very very long text that will definitely require a taller row to display properly when word wrap is enabled.");

            // Enable text wrapping for all populated cells to allow height calculation based on content
            for (int r = 0; r <= cells.MaxDataRow; r++)
            {
                for (int c = 0; c <= cells.MaxDataColumn; c++)
                {
                    Cell cell = cells[r, c];
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        Style style = cell.GetStyle();
                        style.IsTextWrapped = true;
                        cell.SetStyle(style);
                    }
                }
            }

            // Enumerate each row, measure the tallest cell content (in pixels), and set the row height accordingly
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                int maxPixelHeight = 0;

                // Scan all columns in the current row
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell != null && cell.Type != CellValueType.IsNull)
                    {
                        // GetHeightOfValue returns the height needed for the cell's content in pixels
                        int cellPixelHeight = cell.GetHeightOfValue();

                        if (cellPixelHeight > maxPixelHeight)
                            maxPixelHeight = cellPixelHeight;
                    }
                }

                // If any cell required height, apply it to the row
                if (maxPixelHeight > 0)
                {
                    // SetRowHeightPixel sets the row height directly in pixels
                    cells.SetRowHeightPixel(row, maxPixelHeight);
                }
            }

            // Save the workbook to a file
            workbook.Save("DynamicRowHeight.xlsx");
        }
    }
}
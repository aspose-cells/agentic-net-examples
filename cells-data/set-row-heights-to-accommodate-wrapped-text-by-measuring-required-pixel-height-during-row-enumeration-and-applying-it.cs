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

            // Sample data with long text that requires wrapping
            string longText = "This is a very long piece of text that will be wrapped inside the cell. " +
                              "It contains multiple sentences to ensure that the row height needs to be increased " +
                              "to display all the content properly when text wrapping is enabled.";

            // Populate a few rows with wrapped text
            for (int row = 0; row < 5; row++)
            {
                // Put the same long text in column A for demonstration
                cells[row, 0].PutValue(longText);

                // Enable text wrapping for the cell
                Style style = cells[row, 0].GetStyle();
                style.IsTextWrapped = true;
                cells[row, 0].SetStyle(style);
            }

            // Enumerate each row that contains data
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
            {
                int requiredHeightPixel = 0;

                // Examine each cell in the current row to find the maximum required height
                for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                {
                    Cell cell = cells[rowIndex, colIndex];
                    if (cell != null && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        // Get the height needed for the cell's content (in pixels)
                        int cellHeightPixel = cell.GetHeightOfValue();

                        // Keep the greatest height among cells in the row
                        if (cellHeightPixel > requiredHeightPixel)
                            requiredHeightPixel = cellHeightPixel;
                    }
                }

                // If any cell required a height greater than the default, apply it to the row
                if (requiredHeightPixel > 0)
                {
                    // Set row height directly in pixels
                    cells.SetRowHeightPixel(rowIndex, requiredHeightPixel);
                }
            }

            // Save the workbook to a file
            workbook.Save("RowHeightAdjusted.xlsx");
        }
    }
}
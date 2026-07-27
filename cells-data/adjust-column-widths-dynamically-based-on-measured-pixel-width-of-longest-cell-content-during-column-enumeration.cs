using System;
using Aspose.Cells;

namespace DynamicColumnWidthDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Sample data – in real scenarios the worksheet would already
            // contain data that needs column width adjustment.
            // ------------------------------------------------------------
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("A much longer piece of text that should expand the column");
            cells["B1"].PutValue("12345");
            cells["B2"].PutValue("67890");
            cells["C1"].PutValue("Lorem ipsum dolor sit amet, consectetur adipiscing elit.");
            cells["C2"].PutValue("Short");

            // Determine the range that contains data
            int firstRow = 0;
            int lastRow = cells.MaxDataRow;          // last row with data
            int firstColumn = 0;
            int lastColumn = cells.MaxDataColumn;    // last column with data

            // Iterate through each column, autofit based on its content,
            // retrieve the resulting pixel width, and then explicitly set
            // the column width in pixels (adding a small padding for safety).
            for (int col = firstColumn; col <= lastColumn; col++)
            {
                // AutoFitColumn adjusts the column width according to the
                // longest cell content within the specified row range.
                worksheet.AutoFitColumn(col, firstRow, lastRow);

                // Get the width that AutoFitColumn calculated (in pixels)
                int pixelWidth = cells.GetColumnWidthPixel(col);

                // Optional: add a few pixels as padding so text does not touch the border
                int paddedPixelWidth = pixelWidth + 5;

                // Explicitly set the column width in pixels
                cells.SetColumnWidthPixel(col, paddedPixelWidth);
            }

            // Save the workbook
            workbook.Save("DynamicColumnWidthResult.xlsx");
        }
    }
}
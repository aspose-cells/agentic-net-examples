// Title: How to dynamically set Excel row height in C# with Aspose.Cells using the tallest cell’s pixel measurement
// AI Prompts: Generate C# code that loops through each row of an Aspose.Cells worksheet, measures each cell’s pixel height with Cell.GetHeightOfValue, and applies the maximum height to the row using SetRowHeightPixel. | Show how to enable text wrapping for cells, calculate the required row height based on wrapped content, and save the workbook with adjusted row heights in Aspose.Cells. | Provide a complete example that creates a workbook, populates cells with varying text lengths, determines the tallest cell per row, and sets row heights in pixels.
// Common Searches: Aspose.Cells C# set row height based on cell content pixel value | calculate maximum cell height in a row Aspose.Cells .NET | adjust Excel row height automatically for wrapped text using Aspose.Cells | Cell.GetHeightOfValue usage example C# | SetRowHeightPixel per row Aspose.Cells tutorial
// Tags: set row height pixel Aspose.Cells | measure cell pixel height Aspose.Cells | wrap text row height calculation Aspose.Cells | dynamic row height based on content Aspose.Cells | Cell.GetHeightOfValue C# example | iterate rows adjust height Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, adds wrapped text to cells, iterates each row to obtain the pixel height of every non‑null cell via Cell.GetHeightOfValue, selects the maximum height per row, sets the row height with SetRowHeightPixel (including a small padding), and saves the file as DynamicRowHeight.xlsx.
class DynamicRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with varying lengths and enable text wrapping where needed
        cells["A1"].PutValue("Short");
        cells["B1"].PutValue("This is a longer text that should wrap and increase row height.");
        Style wrapStyle = cells["B1"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        cells["B1"].SetStyle(wrapStyle);

        cells["A2"].PutValue("Another row with\nmultiple lines");
        cells["B2"].PutValue("Medium length");
        Style wrapStyle2 = cells["A2"].GetStyle();
        wrapStyle2.IsTextWrapped = true;
        cells["A2"].SetStyle(wrapStyle2);

        // Determine the last row that contains data
        int lastRow = cells.MaxDataRow;

        // Iterate through each row to calculate the tallest cell content
        for (int row = 0; row <= lastRow; row++)
        {
            int maxPixelHeight = 0;

            // Iterate through each column in the current row
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];
                if (cell != null && cell.Type != CellValueType.IsNull)
                {
                    // Get the height of the cell's value in pixels
                    int cellPixelHeight = cell.GetHeightOfValue();

                    // Track the maximum height found in this row
                    if (cellPixelHeight > maxPixelHeight)
                        maxPixelHeight = cellPixelHeight;
                }
            }

            // If any cell contributed a height, set the row height accordingly (add a small padding)
            if (maxPixelHeight > 0)
            {
                cells.SetRowHeightPixel(row, maxPixelHeight + 2);
            }
        }

        // Save the workbook with the dynamically adjusted row heights
        workbook.Save("DynamicRowHeight.xlsx");
    }
}

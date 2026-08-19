// Title: Auto‑adjust row height in Aspose.Cells for .NET using GetHeightOfValue (pixel precision)
// Description: Creates a workbook, adds wrapped text, determines the data range, iterates each row, computes the pixel height needed for every non‑empty cell with GetHeightOfValue, picks the maximum per row, and applies it via SetRowHeightPixel before saving.
// Keywords: Aspose.Cells | C# row height | SetRowHeightPixel | GetHeightOfValue | auto row height .NET | wrap text Excel | dynamic row sizing | pixel height calculation | Excel export Aspose | adjust row height programmatically
// Common Searches: How to auto size row height in Aspose.Cells | GetHeightOfValue example C# | SetRowHeightPixel per row Aspose | Adjust row height for wrapped text .NET | Calculate cell height in pixels Aspose.Cells
// Developer Intent: Programmatically set each worksheet row’s height to the pixel value required by the tallest cell in that row.
// Use Cases: Create Excel reports with multi‑line descriptions that need full visibility. | Export invoices or catalogs where notes vary in length and rows must expand automatically. | Generate data dashboards in Excel with fixed column widths while rows adapt to content.
// AI Prompts: Provide a C# snippet that iterates through all rows in an Aspose.Cells worksheet, uses Cell.GetHeightOfValue for each cell, and applies Worksheet.Cells.SetRowHeightPixel with the maximum height per row. | Show how to enable text wrapping, define column widths, and then automatically resize rows based on the tallest cell using Aspose.Cells for .NET. | Explain the algorithm behind GetHeightOfValue and how it integrates with SetRowHeightPixel to achieve pixel‑accurate row auto‑sizing.

using System;
using Aspose.Cells;

// Creates a workbook, adds wrapped text, determines the data range, iterates each row, computes the pixel height needed for every non‑empty cell with GetHeightOfValue, picks the maximum per row, and applies it via SetRowHeightPixel before saving.
class DynamicRowHeight
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // -------------------------------------------------
        // Sample data with varying lengths and wrapping
        // -------------------------------------------------
        cells["A1"].PutValue("Short");
        cells["B1"].PutValue("This is a longer text that should wrap into multiple lines when column width is limited.");
        Style wrapStyle = cells["B1"].GetStyle();
        wrapStyle.IsTextWrapped = true;
        cells["B1"].SetStyle(wrapStyle);
        cells.SetColumnWidth(1, 20); // Set column B width to force wrapping

        cells["A2"].PutValue("Another row with\nline break");
        cells["B2"].PutValue("Medium length text");
        Style wrapStyleA2 = cells["A2"].GetStyle();
        wrapStyleA2.IsTextWrapped = true;
        cells["A2"].SetStyle(wrapStyleA2);

        // -------------------------------------------------
        // Determine the range that contains data
        // -------------------------------------------------
        int maxRow = cells.MaxDataRow;      // Last row index with data
        int maxCol = cells.MaxDataColumn;   // Last column index with data

        // -------------------------------------------------
        // Enumerate each row and set its height based on the tallest cell content
        // -------------------------------------------------
        for (int row = 0; row <= maxRow; row++)
        {
            int maxPixelHeight = 0;

            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell != null && cell.Type != CellValueType.IsNull)
                {
                    // GetHeightOfValue returns the height needed for the cell's value in pixels
                    int cellHeight = cell.GetHeightOfValue();

                    if (cellHeight > maxPixelHeight)
                        maxPixelHeight = cellHeight;
                }
            }

            // Apply the calculated height to the row (if any cell contributed a height)
            if (maxPixelHeight > 0)
            {
                cells.SetRowHeightPixel(row, maxPixelHeight);
            }
        }

        // -------------------------------------------------
        // Save the workbook
        // -------------------------------------------------
        workbook.Save("DynamicRowHeight.xlsx");
    }
}

// Title: Auto‑Fit Row Height for Wrapped Text in Aspose.Cells for .NET (C#) Using GetHeightOfValue
// Description: Demonstrates how to enable text wrapping, measure the required pixel height of each cell with Cell.GetHeightOfValue, compute the maximum height per row, and apply it via Worksheet.Cells.SetRowHeightPixel for both new and existing workbooks.
// Keywords: Aspose.Cells row height | C# auto fit row height | wrapped text pixel height | GetHeightOfValue example | SetRowHeightPixel .NET | Excel row auto‑size Aspose | measure cell height Aspose.Cells | auto adjust row height C#
// Common Searches: Aspose.Cells auto fit row height wrapped text | GetHeightOfValue set row height pixel | C# adjust Excel row height based on content | Aspose.Cells SetRowHeightPixel usage | how to auto‑size rows in Aspose.Cells
// Developer Intent: Programmatically resize each worksheet row so that wrapped text fits without truncation by calculating the needed pixel height and setting the row height accordingly.
// Use Cases: Generating reports with description fields that may span multiple lines. | Creating invoices where product notes require dynamic row expansion. | Exporting user comments or logs to Excel while preserving multiline formatting.
// AI Prompts: Write a C# method that loops through all rows in an Aspose.Cells worksheet, finds the maximum pixel height of wrapped cells using GetHeightOfValue, and sets the row height with SetRowHeightPixel. | Provide a complete example that loads an existing workbook, enables text wrapping for a specific range, and automatically adjusts row heights based on the cell contents. | Explain the algorithm behind GetHeightOfValue, how it returns pixel height, and how to convert that value to points for point‑based row height settings.

using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightAdjustment
{
    // Demonstrates how to enable text wrapping, measure the required pixel height of each cell with Cell.GetHeightOfValue, compute the maximum height per row, and apply it via Worksheet.Cells.SetRowHeightPixel for both new and existing workbooks.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // Replace with new Workbook("input.xlsx") to load
            Worksheet worksheet = workbook.Worksheets[0];

            // Example data with wrapped text
            worksheet.Cells["A1"].PutValue("This is a long text that should wrap inside the cell and cause the row height to increase.");
            worksheet.Cells["B1"].PutValue("Short text");
            worksheet.Cells["A2"].PutValue("Another long piece of text that will be wrapped and needs more height.");
            worksheet.Cells["B2"].PutValue("More short text");

            // Enable text wrapping for the cells that need it
            for (int r = 0; r <= worksheet.Cells.MaxDataRow; r++)
            {
                for (int c = 0; c <= worksheet.Cells.MaxDataColumn; c++)
                {
                    Cell cell = worksheet.Cells[r, c];
                    if (cell != null && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        Style style = cell.GetStyle();
                        style.IsTextWrapped = true;
                        cell.SetStyle(style);
                    }
                }
            }

            // Enumerate each row, measure the required pixel height, and apply it
            for (int rowIndex = 0; rowIndex <= worksheet.Cells.MaxDataRow; rowIndex++)
            {
                int maxPixelHeight = 0;

                // Scan all cells in the current row
                for (int colIndex = 0; colIndex <= worksheet.Cells.MaxDataColumn; colIndex++)
                {
                    Cell cell = worksheet.Cells[rowIndex, colIndex];
                    if (cell == null) continue;

                    // Only consider cells with text wrapping enabled
                    Style style = cell.GetStyle();
                    if (style.IsTextWrapped)
                    {
                        // Get the height required to display the cell's value (in pixels)
                        int cellPixelHeight = cell.GetHeightOfValue();

                        // Keep the maximum height found in the row
                        if (cellPixelHeight > maxPixelHeight)
                            maxPixelHeight = cellPixelHeight;
                    }
                }

                // If any wrapped cell was found, set the row height accordingly
                if (maxPixelHeight > 0)
                {
                    // SetRowHeightPixel sets the height in pixels directly
                    worksheet.Cells.SetRowHeightPixel(rowIndex, maxPixelHeight);
                }
            }

            // Save the workbook
            workbook.Save("RowHeightAdjusted.xlsx");
        }
    }
}

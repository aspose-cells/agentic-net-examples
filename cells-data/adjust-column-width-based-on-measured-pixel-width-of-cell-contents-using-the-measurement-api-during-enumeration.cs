// Title: Auto‑fit Excel column widths by pixel measurement with Aspose.Cells for .NET
// Description: Demonstrates iterating through populated cells, using Cell.GetWidthOfValue() to obtain each value's pixel width, computing the maximum width per column, adding optional padding, and setting the column width in pixels via Cells.SetColumnWidthPixel() before saving the workbook.
// Keywords: Aspose.Cells set column width pixel | Cell.GetWidthOfValue | auto fit column width .NET | measure cell content pixel Aspose | adjust column width programmatically | Excel column width pixels C# | Aspose.Cells column autosize | pixel based column sizing
// Common Searches: Aspose.Cells set column width in pixels | How to auto fit column width by pixel Aspose | GetWidthOfValue example C# | Calculate max pixel width per column Aspose.Cells | SetColumnWidthPixel usage
// Developer Intent: Programmatically size each worksheet column to the widest cell content measured in pixels, optionally adding padding for readability.
// Use Cases: Export reports where text length varies and columns must prevent truncation. | Create financial spreadsheets where the largest numeric value determines column width. | Generate templates that maintain consistent visual layout across different data sets. | Build automated Excel files for web applications that need precise pixel‑perfect column sizing.
// AI Prompts: Modify the example to factor in cell font size and style when calculating pixel width. | Add logging that outputs the pixel width calculated for each column before applying it. | Show how to convert the pixel width returned by GetWidthOfValue to Excel's column width units. | Explain how to apply the same pixel‑based column sizing to multiple worksheets in a workbook. | Provide a version that respects merged cells while determining column width.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Demonstrates iterating through populated cells, using Cell.GetWidthOfValue() to obtain each value's pixel width, computing the maximum width per column, adding optional padding, and setting the column width in pixels via Cells.SetColumnWidthPixel() before saving the workbook.
public class AdjustColumnWidthByPixel
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data placed in different columns
            cells["A1"].PutValue("Short");
            cells["A2"].PutValue("This is a longer piece of text");
            cells["A3"].PutValue("Very very long text that needs more column width");

            cells["B1"].PutValue("12345");
            cells["B2"].PutValue("67890");
            cells["B3"].PutValue("1234567890");

            cells["C1"].PutValue("Alpha");
            cells["C2"].PutValue("Beta Gamma Delta");
            cells["C3"].PutValue("Omega");

            // Determine the maximum pixel width required for each column
            Dictionary<int, int> maxPixelWidth = new Dictionary<int, int>();

            int maxRow = cells.MaxDataRow;      // last row that contains data
            int maxCol = cells.MaxDataColumn;   // last column that contains data

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    // Skip empty cells
                    if (cell == null || cell.Value == null)
                        continue;

                    // Get the pixel width of the cell's value
                    int valuePixelWidth = cell.GetWidthOfValue();

                    // Track the maximum width per column
                    if (!maxPixelWidth.ContainsKey(col) || valuePixelWidth > maxPixelWidth[col])
                    {
                        maxPixelWidth[col] = valuePixelWidth;
                    }
                }
            }

            // Apply the calculated widths (add a small padding for visual comfort)
            const int paddingPixels = 5;
            foreach (KeyValuePair<int, int> entry in maxPixelWidth)
            {
                int columnIndex = entry.Key;
                int requiredPixelWidth = entry.Value + paddingPixels;

                // Set the column width in pixels using the provided API
                cells.SetColumnWidthPixel(columnIndex, requiredPixelWidth);
            }

            // Save the workbook
            string outputPath = "AdjustedColumnWidths.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        AdjustColumnWidthByPixel.Run();
    }
}

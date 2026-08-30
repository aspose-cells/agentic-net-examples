// Title: Automatically adjust Excel column widths to content pixel size using Aspose.Cells GetWidthOfValue and SetColumnWidthPixel in C#
// AI Prompts: Write a C# method that accepts a Worksheet, loops through each column, uses Cell.GetWidthOfValue to find the widest cell, adds a configurable pixel padding, and applies Cells.SetColumnWidthPixel for precise column sizing. | Generate C# code that enumerates the used rows and columns of a workbook, computes the maximum pixel width per column with GetWidthOfValue, and sets the column width via SetColumnWidthPixel without invoking AutoFitColumn. | Create a reusable C# utility class for Aspose.Cells that provides an AdjustColumnWidthsByPixel(Worksheet ws, int padding) function implementing pixel‑based column auto‑fit logic.
// Common Searches: Aspose.Cells C# set column width in pixels based on cell content | How to use GetWidthOfValue for column auto‑fit in Aspose.Cells | C# calculate maximum pixel width of a column in an Excel workbook using Aspose.Cells | Programmatically adjust Excel column widths with Aspose.Cells measurement API | Difference between SetColumnWidthPixel and AutoFitColumn in Aspose.Cells C#
// Tags: Aspose.Cells column width pixel measurement | C# GetWidthOfValue usage | SetColumnWidthPixel column auto fit | calculate maximum cell pixel width Aspose.Cells | programmatic column resizing Excel C#

using System;
using System.IO;
using Aspose.Cells;

namespace AdjustColumnWidthDemo
{
    // The example creates a workbook, populates cells with varied text and numbers, determines the used range, and for each column measures the widest cell using Cell.GetWidthOfValue. After adding a small pixel padding, it sets the column width with Cells.SetColumnWidthPixel, producing an Excel file where each column perfectly fits its content.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data with varying lengths
                cells["A1"].PutValue("Short");
                cells["A2"].PutValue("This is a much longer piece of text that requires more width");
                cells["B1"].PutValue(12345);
                cells["B2"].PutValue(678901234);
                cells["C1"].PutValue("Very very long text that will need a wide column");
                cells["C2"].PutValue("Mid length");

                // Determine the range of used rows and columns
                int maxRow = cells.MaxDataRow;          // zero‑based index of the last used row
                int maxColumn = cells.MaxDataColumn;    // zero‑based index of the last used column

                // Iterate through each column to find the maximum pixel width of its cells
                for (int col = 0; col <= maxColumn; col++)
                {
                    int maxPixelWidth = 0;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        Cell cell = cells[row, col];

                        // Check if the cell contains a value (not null/empty)
                        if (cell.Type != CellValueType.IsNull)
                        {
                            // Get the pixel width of the cell's displayed value
                            int pixelWidth = cell.GetWidthOfValue();

                            // Keep the largest width found in this column
                            if (pixelWidth > maxPixelWidth)
                                maxPixelWidth = pixelWidth;
                        }
                    }

                    // Add a small padding (e.g., 5 pixels) to avoid clipping
                    int finalPixelWidth = maxPixelWidth + 5;

                    // Set the column width in pixels for normal view
                    cells.SetColumnWidthPixel(col, finalPixelWidth);
                }

                // Define output file name
                string outputPath = "AdjustedColumnWidths.xlsx";

                // Save the workbook with adjusted column widths
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

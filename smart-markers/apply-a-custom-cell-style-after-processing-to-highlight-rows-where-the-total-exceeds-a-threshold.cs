using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightRows
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: fill rows with numeric values in columns A, B, C
            int totalRows = 10;
            for (int r = 0; r < totalRows; r++)
            {
                cells[r, 0].PutValue(r + 10);   // Column A
                cells[r, 1].PutValue((r + 1) * 5); // Column B
                cells[r, 2].PutValue(r * 3);   // Column C
            }

            // Define the threshold for the row total
            double threshold = 100.0;

            // Create a custom style to highlight rows
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.ForegroundColor = Color.Yellow; // Background color
            highlightStyle.Font.Color = Color.Red;         // Font color
            highlightStyle.Font.IsBold = true;

            // Define which style attributes should be applied
            StyleFlag flag = new StyleFlag();
            flag.All = true; // Apply all formatting defined in the style

            // Iterate through each row, calculate the sum of the first three columns,
            // and apply the highlight style if the sum exceeds the threshold
            for (int r = 0; r < totalRows; r++)
            {
                double rowSum = 0;
                for (int c = 0; c < 3; c++)
                {
                    // Ensure the cell contains a numeric value before adding
                    if (cells[r, c].IsNumericValue)
                    {
                        rowSum += cells[r, c].DoubleValue;
                    }
                }

                if (rowSum > threshold)
                {
                    // Apply the style to the entire row
                    cells.ApplyRowStyle(r, highlightStyle, flag);
                }
            }

            // Save the workbook
            workbook.Save("HighlightedRows.xlsx");
        }
    }
}
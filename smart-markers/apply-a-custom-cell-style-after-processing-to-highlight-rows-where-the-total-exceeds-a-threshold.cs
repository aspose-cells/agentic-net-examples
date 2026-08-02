using System;
using Aspose.Cells;
using System.Drawing;

class HighlightRowsByTotal
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your file)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Threshold for highlighting rows
        double threshold = 100.0;

        // Create a custom style for highlighting
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;          // Background color
        highlightStyle.Pattern = BackgroundType.Solid;          // Apply solid fill
        highlightStyle.Font.IsBold = true;                      // Bold font

        // StyleFlag to apply all style attributes
        StyleFlag flag = new StyleFlag();
        flag.All = true;

        // Determine the used range of rows
        int maxRow = cells.MaxDataRow;

        // Define the columns that contribute to the total (adjust as needed)
        int startCol = 0;
        int endCol = 2;

        // Iterate through each row and calculate the total
        for (int row = 0; row <= maxRow; row++)
        {
            double rowTotal = 0;

            for (int col = startCol; col <= endCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    rowTotal += cell.DoubleValue;
                }
            }

            // If the total exceeds the threshold, apply the custom style to the entire row
            if (rowTotal > threshold)
            {
                cells.ApplyRowStyle(row, highlightStyle, flag);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
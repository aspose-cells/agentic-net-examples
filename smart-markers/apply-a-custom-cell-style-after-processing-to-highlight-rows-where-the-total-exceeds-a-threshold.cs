// Title: Apply a custom highlight style to rows where the total exceeds a threshold using Aspose.Cells for .NET
// Common Searches: c# aspocells highlight entire row when column D value is greater than 1000 | how to apply a custom style to rows based on a numeric threshold in Aspose.Cells | aspocells set row background color after checking total column value | iterate through worksheet rows and style rows that exceed a limit with Aspose.Cells | conditional row formatting in Excel using Aspose.Cells .NET API
// Tags: apply custom row style Aspose.Cells | highlight rows exceeding numeric threshold C# | setstyle for entire row Aspose.Cells | conditional formatting based on cell value Aspose.Cells | iterate rows and apply style Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example loads an Excel workbook, defines a numeric threshold, creates a yellow background with red bold font style, and iterates through each data row. For rows where the total in column D exceeds the threshold, the custom style is applied to every cell in that row before saving the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the threshold for highlighting
        double threshold = 1000.0;

        // Create a custom style for rows that exceed the threshold
        Style highlightStyle = workbook.CreateStyle();
        highlightStyle.ForegroundColor = Color.Yellow;
        highlightStyle.Pattern = BackgroundType.Solid;
        highlightStyle.Font.Color = Color.Red;
        highlightStyle.Font.IsBold = true;

        // Assume totals are in column D (zero‑based index 3) starting from row 2 (zero‑based index 1)
        int totalColumnIndex = 3;
        int startRow = 1;

        // Determine the last row and column that contain data
        int lastRow = cells.MaxDataRow;
        int lastColumn = cells.MaxDataColumn;

        // Iterate through each row, check the total, and apply the style if needed
        for (int row = startRow; row <= lastRow; row++)
        {
            Cell totalCell = cells[row, totalColumnIndex];
            if (totalCell.Value != null && double.TryParse(totalCell.Value.ToString(), out double total))
            {
                if (total > threshold)
                {
                    // Apply the highlight style to every cell in the current row
                    for (int col = 0; col <= lastColumn; col++)
                    {
                        cells[row, col].SetStyle(highlightStyle);
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}

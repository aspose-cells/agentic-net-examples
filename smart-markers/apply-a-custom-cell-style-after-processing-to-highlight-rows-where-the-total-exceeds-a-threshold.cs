// Title: Highlight rows exceeding a sum threshold with a custom style using Aspose.Cells for .NET
// Description: C# example that creates a workbook, populates ten rows with three random numbers each, calculates each row's total, and applies a yellow bold style to any row whose sum is greater than 200. The solution uses Aspose.Cells Style, StyleFlag, and ApplyRowStyle before saving the file as HighlightedRows.xlsx.
// Keywords: Aspose.Cells | .NET | C# | ApplyRowStyle | StyleFlag | row highlighting | conditional formatting | threshold | Excel automation | custom style | background color | bold font
// Common Searches: Aspose.Cells highlight rows based on sum | Apply custom style to entire row C# Aspose.Cells | How to use StyleFlag with ApplyRowStyle | Conditional row formatting in Aspose.Cells .NET | Set background color for rows exceeding a value
// Developer Intent: Apply a custom style to whole rows when their summed cell values exceed a predefined threshold.
// Use Cases: Sales report: flag rows where total sales surpass the target. | Inventory list: emphasize rows with stock value above a limit. | Financial statement: mark expense rows that exceed budgeted amounts. | Project tracking: highlight tasks whose total hours go beyond allocated time.
// AI Prompts: Generate C# code with Aspose.Cells that colors rows red when the sum of columns A‑C exceeds 500. | Show how to read the threshold from a worksheet cell instead of a hard‑coded constant. | Create an example that uses StyleFlag to apply both bold font and a green fill to qualifying rows. | Explain how to extend the code to apply different styles based on multiple threshold ranges.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightRows
{
    // C# example that creates a workbook, populates ten rows with three random numbers each, calculates each row's total, and applies a yellow bold style to any row whose sum is greater than 200. The solution uses Aspose.Cells Style, StyleFlag, and ApplyRowStyle before saving the file as HighlightedRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data: three numeric columns per row
            int rows = 10;
            int cols = 3;
            Random rnd = new Random();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    cells[r, c].PutValue(rnd.Next(10, 100));
                }
            }

            // Define a threshold for the row total
            double threshold = 200;

            // Create a style to highlight rows that exceed the threshold
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.BackgroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;
            highlightStyle.Font.IsBold = true;

            // StyleFlag indicating which parts of the style to apply (apply all)
            StyleFlag flag = new StyleFlag { All = true };

            // Iterate through each row, calculate the sum, and apply the style if needed
            for (int r = 0; r < rows; r++)
            {
                double rowSum = 0;
                for (int c = 0; c < cols; c++)
                {
                    // Use GetValue to retrieve the numeric value
                    rowSum += cells[r, c].DoubleValue;
                }

                if (rowSum > threshold)
                {
                    // Apply the highlight style to the entire row
                    cells.ApplyRowStyle(r, highlightStyle, flag);
                }
            }

            // Save the workbook
            workbook.Save("HighlightedRows.xlsx");
        }
    }
}

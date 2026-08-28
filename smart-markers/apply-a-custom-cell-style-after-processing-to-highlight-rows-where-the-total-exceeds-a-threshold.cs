// Title: Highlight Excel rows with a yellow background when the sum of numeric cells exceeds a threshold using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, computes the total of numeric cells in each row, creates a yellow solid style, and applies it to rows whose sum is greater than a given limit. | Modify an existing Aspose.Cells workbook to add a custom StyleFlag and use ApplyRowStyle to conditionally format rows that surpass a specified numeric total.
// Common Searches: aspocells c# apply yellow background to rows where row total > 100 | conditional row formatting based on sum of cells using Aspose.Cells .NET | how to set background color for rows exceeding a numeric limit in Aspose.Cells | calculate row totals and highlight rows in Excel programmatically with Aspose.Cells | using Aspose.Cells to format rows after aggregating cell values in C#
// Tags: row-level conditional styling Aspose.Cells | ApplyRowStyle with StyleFlag C# | row sum threshold styling .NET | bright yellow row highlight Aspose.Cells | compute numeric row total Aspose.Cells | style rows based on aggregate value C#

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHighlightRows
{
    // The program loads input.xlsx, sums numeric cells in each row, creates a yellow solid background style, applies it to rows whose total exceeds 100, and saves the result as output.xlsx.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the threshold for highlighting
            double threshold = 100.0;

            // Create a style that will be applied to rows exceeding the threshold
            Style highlightStyle = workbook.CreateStyle();
            highlightStyle.BackgroundColor = Color.Yellow;
            highlightStyle.Pattern = BackgroundType.Solid;

            // StyleFlag to apply all style attributes
            StyleFlag flag = new StyleFlag { All = true };

            // Iterate through each used row
            int maxRow = cells.MaxDataRow; // last row with data
            for (int row = 0; row <= maxRow; row++)
            {
                double rowTotal = 0.0;

                // Sum numeric values in the current row
                int maxCol = cells.MaxDataColumn;
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        rowTotal += cell.DoubleValue;
                    }
                }

                // Apply the highlight style if the total exceeds the threshold
                if (rowTotal > threshold)
                {
                    cells.ApplyRowStyle(row, highlightStyle, flag);
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}

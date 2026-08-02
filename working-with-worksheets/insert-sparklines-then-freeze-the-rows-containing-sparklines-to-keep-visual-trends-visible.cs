// Title: Insert Line Sparklines per Row and Freeze Those Rows – Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, fill a 5×5 numeric range, add a line sparkline for each row in column F via a SparklineGroup, freeze rows 0‑4 so the sparklines stay visible while scrolling, and save the file as SparklinesWithFrozenRows.xlsx.
// Keywords: Aspose.Cells C# sparkline | Aspose.Cells FreezePanes | add sparklines Aspose.Cells | freeze rows Aspose.Cells | Excel sparkline group .NET | C# Excel freeze panes | GitHub Aspose.Cells sparkline example | global | US
// Common Searches: C# Aspose.Cells add sparklines to each row | How to freeze rows after inserting sparklines with Aspose.Cells | Aspose.Cells FreezePanes example for sparkline groups | Create line sparklines in Excel using Aspose.Cells .NET | Freeze top rows in workbook with Aspose.Cells
// Developer Intent: Create an Excel workbook, populate data, attach a line sparkline to every row, freeze those rows, and save the result.
// Use Cases: Sales dashboard where each product row shows a trend sparkline and the rows stay visible while scrolling. | Financial report that adds per‑row sparklines for month‑over‑month performance and freezes the rows for quick reference. | Sensor‑data sheet that generates a sparkline for each measurement row and locks the rows to act as a persistent trend pane.
// AI Prompts: Generate C# code using Aspose.Cells to add vertical sparklines for each column and freeze the first three rows. | Modify the sample to use a Column sparkline type instead of Line and keep rows 1‑5 frozen. | Explain how FreezePanes parameters correspond to Excel's Freeze Panes UI when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklinesFreezeDemo
{
    // Shows how to create a workbook, fill a 5×5 numeric range, add a line sparkline for each row in column F via a SparklineGroup, freeze rows 0‑4 so the sparklines stay visible while scrolling, and save the file as SparklinesWithFrozenRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for multiple rows (5 rows, 5 columns)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Define the location range where sparklines will be placed (column F, rows 0-4)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 5,   // Column index 5 = column F
                EndColumn = 5
            };

            // Add a sparkline group for the data range A1:E5 (horizontal layout)
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:E5",
                false,               // isVertical = false (plot by column)
                sparklineLocation);

            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline for each row of data
            for (int r = 0; r < 5; r++)
            {
                // Data range for the current row (e.g., A1:E1, A2:E2, ...)
                string dataRange = $"A{r + 1}:E{r + 1}";
                // Place the sparkline in column F of the same row
                group.Sparklines.Add(dataRange, r, 5);
            }

            // Freeze the rows that contain sparklines (rows 0‑4)
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Use row index 5 (cell after the frozen area) and freeze 5 rows.
            sheet.FreezePanes(5, 0, 5, 0);

            // Save the workbook
            workbook.Save("SparklinesWithFrozenRows.xlsx");
        }
    }
}

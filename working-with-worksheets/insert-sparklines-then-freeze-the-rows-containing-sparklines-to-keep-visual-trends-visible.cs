// Title: Insert row‑wise line sparklines and freeze those rows with Aspose.Cells for .NET
// Description: Creates a workbook, fills a 5×4 data range, adds a line sparkline for each row in column E, freezes the first five rows so the sparklines remain visible while scrolling, and saves the file as SparklinesWithFrozenRows.xlsx.
// Keywords: Aspose.Cells | C# | sparklines | line sparkline | freeze panes | freeze rows | Excel dashboard | SparklineGroup | FreezePanes method | row‑wise sparklines
// Common Searches: Aspose.Cells add sparklines per row | freeze top rows after inserting sparklines .NET | how to keep sparklines visible while scrolling Excel | C# example for SparklineGroup and FreezePanes | insert line sparklines in column E using Aspose.Cells
// Developer Intent: Add a line sparkline for each data row and lock those rows in view.
// Use Cases: Financial statements where each row shows a mini trend chart next to the values. | Excel dashboards that need static trend indicators while users scroll through large datasets. | Automated report generation that embeds row‑level sparklines and keeps them anchored for quick comparison.
// AI Prompts: Generate C# code to create a line sparkline for every row of A1:D5 and place it in column E with Aspose.Cells. | Explain the parameters of Worksheet.FreezePanes to freeze the first N rows after adding sparklines. | Show how to modify the example to use column F for sparklines and also freeze the first column. | Provide a step‑by‑step guide to add sparklines and freeze panes in an Aspose.Cells workbook for a sales dashboard.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills a 5×4 data range, adds a line sparkline for each row in column E, freezes the first five rows so the sparklines remain visible while scrolling, and saves the file as SparklinesWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (5 rows × 4 columns)
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Define the location range where sparklines will be placed (column E, rows 0‑4)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 4,
            StartColumn = 4,   // Column index 4 = column "E"
            EndColumn = 4
        };

        // Add a sparkline group for the data range A1:D5.
        // isVertical = false because we plot by rows.
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D5", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add a sparkline for each row (A:D of that row) into column E of the same row
        for (int r = 0; r < 5; r++)
        {
            string dataRange = $"A{r + 1}:D{r + 1}";
            group.Sparklines.Add(dataRange, r, 4);
        }

        // Freeze the rows that contain the sparklines (rows 0‑4).
        // Freeze first 5 rows, no columns are frozen.
        sheet.FreezePanes(5, 1, 5, 0);

        // Save the workbook
        workbook.Save("SparklinesWithFrozenRows.xlsx");
    }
}

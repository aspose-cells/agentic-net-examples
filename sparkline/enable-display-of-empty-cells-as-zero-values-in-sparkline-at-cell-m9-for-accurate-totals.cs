using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineEmptyCellAsZero
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample data for the sparkline (you can adjust the range as needed)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(null); // empty cell
        sheet.Cells["C1"].PutValue(3);
        sheet.Cells["D1"].PutValue(7);

        // Define the location where the sparkline will be placed: cell M9
        // Column M = 12 (0‑based), Row 9 = 8 (0‑based)
        CellArea location = new CellArea
        {
            StartColumn = 12,
            EndColumn = 12,
            StartRow = 8,
            EndRow = 8
        };

        // Add a sparkline group with the data range A1:D1 and place it at M9
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Ensure empty cells are plotted as zeros for accurate totals
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // (Optional) Add the sparkline to the group – Add method also creates the sparkline
        // The data range is already specified in the Add call above, so this line is not strictly required.
        // group.Sparklines.Add(sheet.Name + "!A1:D1", 8, 12);

        // Save the workbook
        workbook.Save("SparklineWithZeroEmptyCells.xlsx");
    }
}
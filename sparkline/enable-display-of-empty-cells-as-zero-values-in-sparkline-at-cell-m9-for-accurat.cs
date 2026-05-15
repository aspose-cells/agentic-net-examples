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

        // Sample data for the sparkline (some cells may be empty)
        sheet.Cells["A9"].PutValue(5);
        sheet.Cells["B9"].PutValue(null); // empty cell
        sheet.Cells["C9"].PutValue(2);
        sheet.Cells["D9"].PutValue(7);

        // Define the location where the sparkline will be placed (cell M9)
        // Column M -> index 12, Row 9 -> index 8 (zero‑based)
        CellArea location = new CellArea
        {
            StartColumn = 12,
            EndColumn = 12,
            StartRow = 8,
            EndRow = 8
        };

        // Add a sparkline group: line type, data range A9:D9, horizontal layout, location M9
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A9:D9", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Configure the group to treat empty cells as zero values
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // (Optional) Add the sparkline explicitly – the Add method already creates it,
        // but calling Add on the collection ensures the sparkline exists.
        group.Sparklines.Add(sheet.Name + "!A9:D9", 8, 12);

        // Save the workbook
        workbook.Save("SparklineWithZeroForEmpty.xlsx");
    }
}
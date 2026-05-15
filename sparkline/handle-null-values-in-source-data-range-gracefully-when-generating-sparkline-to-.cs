using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineNullHandlingDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the source data range with some empty (null) cells
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(null); // empty cell
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(null); // empty cell
        sheet.Cells["E1"].PutValue(3);

        // Define the location where the sparkline will be placed (single cell F1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group with the data range and location
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Handle null values gracefully:
        // - NotPlotted : leaves a gap
        // - Zero       : treats null as 0
        // - Interpolated: interpolates between surrounding values
        // Choose Interpolated to avoid rendering errors while keeping visual continuity
        group.PlotEmptyCellsType = PlotEmptyCellsType.Interpolated;

        // Save the workbook
        workbook.Save("SparklineNullHandling.xlsx");
    }
}
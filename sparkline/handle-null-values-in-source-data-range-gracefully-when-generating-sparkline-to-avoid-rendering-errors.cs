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

        // Populate the source range with some values and nulls (empty cells)
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(null); // empty cell
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(null); // empty cell
        sheet.Cells["E1"].PutValue(4);

        // Define where the sparkline will be placed (cell F1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group for the range A1:E1 and place it in F1
        int groupIndex = sheet.SparklineGroups.Add(
            SparklineType.Line,   // type of sparkline
            "A1:E1",              // data range containing nulls
            false,                // plot by row (horizontal)
            sparklineLocation);   // location range

        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure how empty cells are handled to avoid rendering errors.
        // Options: NotPlotted (gap), Zero (display as 0), Interpolated (estimate value)
        group.PlotEmptyCellsType = PlotEmptyCellsType.Zero;

        // Save the workbook
        workbook.Save("SparklineNullHandlingDemo.xlsx");
    }
}
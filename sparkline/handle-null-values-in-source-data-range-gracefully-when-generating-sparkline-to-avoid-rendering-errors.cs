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

        // Define the location where the sparkline will be placed (cell F1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 5,
            EndColumn = 5
        };

        // Add a sparkline group for the data range A1:E1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:E1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Configure how empty cells are plotted to avoid rendering errors.
        // Options: NotPlotted (gap), Zero, Interpolated.
        // Here we choose Interpolated to smoothly bridge missing values.
        group.PlotEmptyCellsType = PlotEmptyCellsType.Interpolated;

        // Save the workbook
        workbook.Save("SparklineNullHandlingDemo.xlsx");
    }
}
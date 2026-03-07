using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineAxisOptionsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add date values for the horizontal axis (row 1)
        sheet.Cells["A1"].PutValue(new DateTime(2023, 1, 1));
        sheet.Cells["B1"].PutValue(new DateTime(2023, 2, 1));
        sheet.Cells["C1"].PutValue(new DateTime(2023, 3, 1));
        sheet.Cells["D1"].PutValue(new DateTime(2023, 4, 1));

        // Add numeric data for the sparkline (row 2)
        sheet.Cells["A2"].PutValue(5);
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["C2"].PutValue(1);
        sheet.Cells["D2"].PutValue(3);

        // Define the cell area where the sparkline will be placed (E2)
        CellArea location = new CellArea
        {
            StartRow = 1,
            EndRow = 1,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with the data range A2:D2
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A2:D2", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];

        // Configure horizontal axis options
        group.ShowHorizontalAxis = true;                     // display the horizontal axis
        CellsColor hAxisColor = workbook.CreateCellsColor(); // create a color object
        hAxisColor.Color = Color.Gray;                      // set axis line color
        group.HorizontalAxisColor = hAxisColor;             // apply color
        group.HorizontalAxisDateRange = "A1:D1";             // associate date range with axis

        // Configure vertical axis options
        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group; // same max for all sparklines
        group.VerticalAxisMaxValue = 10.0;                               // custom maximum
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom; // custom minimum
        group.VerticalAxisMinValue = 0.0;                                // set minimum

        // Save the workbook to a file
        workbook.Save("SparklineAxisOptions.xlsx");
    }
}
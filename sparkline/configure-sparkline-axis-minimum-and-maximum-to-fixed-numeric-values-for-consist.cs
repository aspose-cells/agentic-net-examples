using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineAxisFixedScaleDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(8);
        sheet.Cells["D1"].PutValue(3);

        // Define where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group with a line sparkline
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (optional, already added by Add method)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Set fixed minimum and maximum values for the vertical axis
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom; // Use custom min
        group.VerticalAxisMinValue = 0.0;                                 // Fixed minimum

        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Custom; // Use custom max
        group.VerticalAxisMaxValue = 10.0;                                // Fixed maximum

        // Optional: style the sparkline for better visibility
        CellsColor seriesColor = workbook.CreateCellsColor();
        seriesColor.Color = Color.DarkBlue;
        group.SeriesColor = seriesColor;
        group.LineWeight = 1.0;

        // Save the workbook
        workbook.Save("SparklineFixedAxisScale.xlsx");
    }
}
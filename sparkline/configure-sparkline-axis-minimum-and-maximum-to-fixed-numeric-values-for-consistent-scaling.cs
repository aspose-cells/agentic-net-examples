using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

class SparklineAxisFixedScale
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["C1"].PutValue(1);
        sheet.Cells["D1"].PutValue(8);

        // Define the location where the sparkline will be placed (cell E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a sparkline group of type Line, using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (optional, already added by Add method)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // Configure the vertical axis to use custom fixed minimum and maximum values
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Custom;
        group.VerticalAxisMinValue = 0.0;   // Fixed minimum

        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Custom;
        group.VerticalAxisMaxValue = 10.0;  // Fixed maximum

        // Optional: set some visual properties
        group.ShowHighPoint = true;
        group.ShowLowPoint = true;
        group.LineWeight = 1.0;

        // Save the workbook
        workbook.Save("SparklineFixedAxis.xlsx");
    }
}
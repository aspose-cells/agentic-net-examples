using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ConfigureSparklineAxis
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(-2);
        sheet.Cells["C1"].PutValue(3);
        sheet.Cells["D1"].PutValue(0);

        // Define the cell area where the sparkline will be placed (E1)
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line sparkline group using the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Show the horizontal axis and set its color
        group.ShowHorizontalAxis = true;
        CellsColor hAxisColor = workbook.CreateCellsColor();
        hAxisColor.Color = Color.Gray;
        group.HorizontalAxisColor = hAxisColor;

        // (Optional) Set a date range for the horizontal axis
        group.HorizontalAxisDateRange = "A1:D1";

        // Configure vertical axis settings: same max/min for the whole group with custom values
        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
        group.VerticalAxisMaxValue = 10.0;
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;
        group.VerticalAxisMinValue = -5.0;

        // Save the workbook to an XLSX file
        workbook.Save("SparklineAxisConfigured.xlsx");
    }
}
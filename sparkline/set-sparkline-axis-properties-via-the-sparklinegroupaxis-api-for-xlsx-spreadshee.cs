using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

class SparklineAxisDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the sparkline (values 10, 30, 20, 40)
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(30);
        sheet.Cells["C1"].PutValue(20);
        sheet.Cells["D1"].PutValue(40);

        // Define the location where the sparkline will be placed (cell E1)
        CellArea sparklineLocation = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4, // Column E (0‑based index)
            EndColumn = 4
        };

        // Add a line sparkline group that uses the data range A1:D1
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Add the sparkline to the group (required when using Add with dataRange)
        group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

        // -------------------- Axis Settings --------------------
        // Use the same vertical axis scaling for all sparklines in the group
        group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
        group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;

        // Define custom min/max values to ensure consistent scaling
        group.VerticalAxisMaxValue = 50.0; // Upper bound
        group.VerticalAxisMinValue = 0.0;  // Lower bound

        // Show the horizontal axis (useful when data crosses zero)
        group.ShowHorizontalAxis = true;

        // Set the color of the horizontal axis to a dark gray
        CellsColor axisColor = workbook.CreateCellsColor();
        axisColor.Color = Color.DarkGray;
        group.HorizontalAxisColor = axisColor;

        // (Optional) If the data represents dates, specify the date range for the horizontal axis
        // group.HorizontalAxisDateRange = "A1:D1";

        // -------------------------------------------------------

        // Save the workbook to an XLSX file
        workbook.Save("SparklineAxisDemo.xlsx");
    }
}
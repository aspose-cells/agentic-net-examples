using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RotateYaxisTickLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the primary Y axis (ValueAxis) tick labels
        // Disable automatic rotation and set the rotation angle to 90 degrees clockwise
        chart.ValueAxis.TickLabels.IsAutomaticRotation = false;
        chart.ValueAxis.TickLabels.RotationAngle = 90;

        // Save the workbook
        workbook.Save("YaxisTickLabelsRotated.xlsx");
    }
}
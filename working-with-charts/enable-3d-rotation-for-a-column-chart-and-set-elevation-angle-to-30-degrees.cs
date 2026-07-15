using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

public class Enable3DRotationDemo
{
    public static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a 3‑D column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the chart data range
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable 3‑D rotation (rotation around the Z‑axis)
        chart.RotationAngle = 30; // valid range 0‑44 for Column3D

        // Set elevation angle to 30 degrees
        chart.Elevation = 30;

        // Save the workbook
        workbook.Save("Column3D_Rotation_Elevation.xlsx");
    }
}
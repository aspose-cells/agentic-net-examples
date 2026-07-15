using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCells3DPerspectiveDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["B4"].PutValue(3000);

            // Add a 3‑D column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable perspective projection
            chart.RightAngleAxes = false;

            // Apply 3‑D perspective and rotation settings
            chart.Perspective = 30;      // Perspective angle (0‑100)
            chart.RotationAngle = 45;    // Rotation around the Z‑axis (0‑360)
            chart.DepthPercent = 150;    // Depth of the chart as a percentage of width (20‑2000)

            // Save the workbook (lifecycle: save)
            workbook.Save("Column3DPerspective.xlsx");

            Console.WriteLine("3‑D column chart with perspective applied successfully.");
        }
    }
}
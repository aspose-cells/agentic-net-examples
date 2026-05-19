using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChart3DRotationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
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

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set three‑dimensional rotation angles
            // X‑axis rotation (elevation) = 20 degrees
            chart.Elevation = 20;

            // Y‑axis rotation – Aspose.Cells does not expose a direct Y‑axis rotation for charts.
            // The closest property affecting the Y‑axis view is Perspective.
            // Here we set Perspective to 45 degrees to simulate a Y‑axis rotation effect.
            chart.Perspective = 45;

            // Z‑axis rotation (around the Z‑axis) = 10 degrees
            chart.RotationAngle = 10;

            // Save the workbook with the configured chart
            workbook.Save("Chart3DRotationDemo.xlsx");

            Console.WriteLine("Chart created with X=20, Y=45 (Perspective), Z=10 rotation angles.");
        }
    }
}
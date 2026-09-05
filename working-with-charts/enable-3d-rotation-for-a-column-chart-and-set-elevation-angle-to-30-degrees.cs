// Title: Enable 3‑D rotation and set a 30° elevation for a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds sample data, inserts a 3‑D column chart, and sets RotationAngle to 45° and Elevation to 30°. | Show how to modify the 3‑D view of a column chart by configuring its RotationAngle and Elevation properties in Aspose.Cells for .NET. | Provide a complete example that saves an XLSX file containing a 3‑D column chart with custom rotation and elevation settings.
// Common Searches: Aspose.Cells C# set rotation angle for 3D column chart | How to change elevation of a 3D chart using Aspose.Cells .NET | C# example configuring 3D view of column chart with Aspose.Cells | Set 3D rotation and elevation for column chart in Aspose.Cells | Aspose.Cells 3D column chart customization tutorial
// Tags: Aspose.Cells set 3D chart rotation angle | Aspose.Cells column chart elevation property | Aspose.Cells configure 3D view .NET | Aspose.Cells create 3D column chart C# | Aspose.Cells save workbook with 3D chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCells3DRotationDemo
{
    // Demonstrates creating a workbook, adding data, inserting a 3‑D column chart, setting RotationAngle to 45° and Elevation to 30°, and saving the file as an XLSX using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
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

            // Enable 3‑D rotation by setting the RotationAngle (0‑360 degrees)
            chart.RotationAngle = 45; // example rotation

            // Set the elevation angle to 30 degrees
            chart.Elevation = 30;

            // Save the workbook
            workbook.Save("3DRotationColumnChart.xlsx");

            Console.WriteLine("3‑D column chart created with rotation and elevation set.");
        }
    }
}

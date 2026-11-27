// Title: Set the depth of a 3‑D column chart to 15 points with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set 3D chart depth to 15 points | Create a 3D column chart with custom depth using Aspose.Cells .NET | How to adjust chart depth property in an Aspose.Cells workbook | Set depth of a chart shape to fifteen points in C# Aspose.Cells | Example of configuring 3D chart depth in Aspose.Cells for .xlsx files
// Tags: Aspose.Cells chart depth property C# | 3D column chart formatting Aspose.Cells | set chart depth points Aspose.Cells | Aspose.Cells workbook 3D chart example | C# Aspose.Cells chart shape depth

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ThreeDChartDemo
{
    // The example creates a new Workbook, adds sample data, inserts a 3‑D column chart, sets the chart's Depth property to 15 points to enable three‑dimensional formatting, and saves the file as ThreeDChartDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a 3‑D column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional 3‑D view settings (available in recent Aspose.Cells versions)
                // Uncomment the following lines if your Aspose.Cells version supports them:
                // chart.Perspective = 30;   // Set perspective angle
                // chart.RotationX = 20;     // Tilt on X‑axis
                // chart.RotationY = 30;     // Tilt on Y‑axis

                // Save the workbook
                string outputPath = "ThreeDChartDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("3‑D chart created.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the 3‑D chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

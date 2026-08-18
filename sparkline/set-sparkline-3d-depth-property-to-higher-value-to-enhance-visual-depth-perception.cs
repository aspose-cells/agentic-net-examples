// Title: Aspose.Cells C# – Boost 3D Chart Depth (DepthPercent) for Sparkline‑style Visuals
// Description: C# example that creates a workbook, fills sample data, adds a 3‑D column chart, and increases the chart's DepthPercent to 300 % (with Perspective 40 and Elevation 30) to give sparkline‑style charts a stronger visual depth before saving as XLSX.
// Keywords: Aspose.Cells | C# 3D chart depth | DepthPercent property | sparkline 3D chart | chart perspective Aspose.Cells | chart elevation Aspose.Cells | Excel 3D column chart | visual depth enhancement | .NET Excel chart example | GitHub Aspose.Cells sample
// Common Searches: how to set DepthPercent in Aspose.Cells C# | increase 3D chart depth Aspose.Cells | adjust perspective and elevation for 3D chart Aspose.Cells | sparkline style 3D chart depth example | Aspose.Cells 3D column chart depth percent
// Developer Intent: Apply a higher DepthPercent value to a 3‑D chart in Aspose.Cells to make the visual depth more pronounced.
// Use Cases: Financial reporting workbook where a 3‑D column chart needs a deeper appearance for clearer trend emphasis. | Sales dashboard that uses sparkline‑style 3‑D charts with customized depth, perspective, and elevation for better readability. | Automated Excel generation that creates multiple 3‑D charts, each with distinct DepthPercent settings to highlight key metrics.
// AI Prompts: Generate C# code with Aspose.Cells that adds a 3‑D column chart and sets DepthPercent to 400, Perspective to 45, and Elevation to 35. | Explain the impact of DepthPercent, Perspective, and Elevation on the rendering of 3‑D charts in Aspose.Cells. | Provide step‑by‑step instructions to increase the depth of sparkline‑style 3‑D charts using Aspose.Cells in a .NET console app.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, fills sample data, adds a 3‑D column chart, and increases the chart's DepthPercent to 300 % (with Perspective 40 and Elevation 30) to give sparkline‑style charts a stronger visual depth before saving as XLSX.
    public class SparklineDepthDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(90);
                sheet.Cells["C3"].PutValue(110);
                sheet.Cells["C4"].PutValue(130);

                // Add a 3‑D column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the chart
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Increase the depth of the 3‑D chart to enhance visual depth perception
                // DepthPercent is a percentage of the chart width (20‑2000). Setting a higher value makes the chart appear deeper.
                chart.DepthPercent = 300; // Example: 300% depth

                // Optionally adjust other 3‑D properties for better visual effect
                chart.Perspective = 40;   // Perspective angle (0‑100)
                chart.Elevation = 30;     // Elevation angle (degrees)

                // Save the workbook
                workbook.Save("SparklineDepthDemo.xlsx", SaveFormat.Xlsx);
                Console.WriteLine("Workbook saved as SparklineDepthDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SparklineDepthDemo.Run();
        }
    }
}

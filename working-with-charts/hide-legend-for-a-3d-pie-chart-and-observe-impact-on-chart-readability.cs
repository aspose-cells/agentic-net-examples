// Title: Hide the legend of a 3‑D Pie chart in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide legend on 3D pie chart | remove legend from Pie3D chart programmatically with Aspose.Cells | Chart.ShowLegend false example for Excel 3D pie chart in .NET | how to disable chart legend in Excel using Aspose.Cells API | C# code to create 3D pie chart without legend using Aspose.Cells
// Tags: Aspose.Cells hide chart legend | Chart.ShowLegend property .NET | 3D pie chart legend removal Aspose.Cells | Excel chart customization C# Aspose | Pie3D chart formatting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace HideLegend3DPieChart
{
    // The sample creates a new workbook, populates cells A1:B5 with fruit categories and values, adds a 3‑D Pie chart (Pie3D) to the worksheet, sets the data range including headers, disables the legend by setting Chart.ShowLegend to false, adds a title, and saves the file as HideLegend3DPieChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["A5"].PutValue("Dates");
                sheet.Cells["B5"].PutValue(25);

                // Add a 3‑D pie chart (Pie3D) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pie3D, 7, 0, 22, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including headers)
                chart.SetChartDataRange("A1:B5", true);

                // Hide the legend using the Chart.ShowLegend property
                chart.ShowLegend = false;

                // Optional: give the chart a title
                chart.Title.Text = "Fruit Distribution (3‑D Pie) – Legend Hidden";

                // Save the workbook to a file
                string outputPath = "HideLegend3DPieChart.xlsx";
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("The 3‑D pie chart has its legend hidden. Open the file in Excel to observe the impact on readability.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

// Title: C# – Remove Gridlines from a Pie Chart using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, inserts a pie chart, and disables major/minor gridlines on both axes plus the 2‑D walls/gridlines property, then saves as PieChart_NoGridlines.xlsx.
// Keywords: Aspose.Cells pie chart gridlines | hide chart gridlines C# | disable axes gridlines Aspose.Cells | WallsAndGridlines2D false | .NET chart formatting | remove pie chart walls
// Common Searches: Aspose.Cells hide gridlines on pie chart | remove chart walls and gridlines .NET | disable major and minor gridlines Aspose.Cells | C# pie chart without gridlines Aspose | how to turn off chart walls Aspose.Cells
// Developer Intent: Generate a pie chart and eliminate every visual gridline and wall to produce a clean, minimalist chart.
// Use Cases: Prepare a sales‑distribution pie chart for presentations without distracting gridlines. | Design a financial dashboard that displays multiple minimalist pie charts. | Export printable reports where pie charts must appear uncluttered.
// AI Prompts: Write C# code that hides all gridlines for any Aspose.Cells chart, including the WallsAndGridlines2D setting. | Explain how the WallsAndGridlines2D property affects pie chart rendering in Aspose.Cells. | Suggest alternative techniques to remove visual gridlines from a pie chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a pie chart, and disables major/minor gridlines on both axes plus the 2‑D walls/gridlines property, then saves as PieChart_NoGridlines.xlsx.
    public class RemovePieChartGridlines
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(20);

                // Add a pie chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide gridlines (pie charts have no axes, but objects exist)
                chart.ValueAxis.MajorGridLines.IsVisible = false;
                chart.ValueAxis.MinorGridLines.IsVisible = false;
                chart.CategoryAxis.MajorGridLines.IsVisible = false;
                chart.CategoryAxis.MinorGridLines.IsVisible = false;

                // Ensure 2‑D gridlines are disabled (safe for pie charts)
                chart.WallsAndGridlines2D = false;

                // Save the workbook
                workbook.Save("PieChart_NoGridlines.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RemovePieChartGridlines.Run();
        }
    }
}

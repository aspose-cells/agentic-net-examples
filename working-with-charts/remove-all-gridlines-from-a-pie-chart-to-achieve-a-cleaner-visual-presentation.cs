// Title: C# – Remove all gridlines from a pie chart with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a pie chart, and disables major and minor gridlines on both value and category axes for a clean visual, then saves the file as PieChart_NoGridlines.xlsx.
// Keywords: Aspose.Cells | C# | .NET | pie chart | gridlines | hide chart gridlines | disable major gridlines | disable minor gridlines | Excel chart formatting | Aspose.Cells chart API
// Common Searches: Aspose.Cells hide pie chart gridlines C# | remove gridlines from chart axes Aspose.Cells | disable major and minor gridlines in Excel chart .NET | how to turn off chart gridlines using Aspose.Cells | C# code to create a pie chart without gridlines
// Developer Intent: Programmatically turn off every gridline on a pie chart to produce a cleaner Excel presentation.
// Use Cases: Generate a sales‑distribution pie chart for a slide deck without any axis gridlines. | Standardize corporate dashboards so all charts share a grid‑line‑free style. | Automate monthly reports that automatically hide chart gridlines before distribution.
// AI Prompts: Write C# code with Aspose.Cells that creates a pie chart and hides both major and minor gridlines on its axes. | Explain how to toggle gridline visibility for different chart types using Aspose.Cells for .NET. | Provide a step‑by‑step guide to verify that gridlines are hidden after saving an Excel workbook containing a pie chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a pie chart, and disables major and minor gridlines on both value and category axes for a clean visual, then saves the file as PieChart_NoGridlines.xlsx.
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

                // Disable all gridlines for a cleaner appearance
                chart.ValueAxis.MajorGridLines.IsVisible = false;
                chart.ValueAxis.MinorGridLines.IsVisible = false;
                chart.CategoryAxis.MajorGridLines.IsVisible = false;
                chart.CategoryAxis.MinorGridLines.IsVisible = false;

                // Save the workbook
                string outputPath = "PieChart_NoGridlines.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemovePieChartGridlines.Run();
        }
    }
}

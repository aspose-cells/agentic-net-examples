// Title: Hide Legend in a Doughnut Chart and Auto‑Expand the Chart Area – Aspose.Cells for .NET
// Description: C# example that creates a workbook, adds category/value data, inserts a doughnut chart, disables the legend with chart.ShowLegend = false, and saves the file. The chart area automatically expands to fill the space left by the hidden legend.
// Keywords: Aspose.Cells hide legend | doughnut chart legend .NET | chart.ShowLegend false | expand chart area Aspose.Cells | C# Excel chart customization | Aspose.Cells doughnut example
// Common Searches: how to hide legend in doughnut chart Aspose.Cells | chart area expand after removing legend Aspose.Cells | Aspose.Cells chart.ShowLegend property usage | C# create doughnut chart without legend | Aspose.Cells example hide chart legend
// Developer Intent: Remove the legend from a doughnut chart so the chart area automatically occupies the freed space.
// Use Cases: Generate compact Excel reports where a legend is unnecessary, maximizing visual space. | Build dashboard widgets that display doughnut charts without legends for a cleaner look. | Automate workbook creation for data summaries, ensuring charts fill their containers after the legend is hidden.
// AI Prompts: Provide C# code using Aspose.Cells to add a doughnut chart, hide its legend, and verify that the chart area expands to use the freed space. | Show how to check the value of chart.ShowLegend after setting it to false in an Aspose.Cells doughnut chart. | Create an Aspose.Cells .NET example that builds a doughnut chart without a legend and saves the workbook as an .xlsx file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds category/value data, inserts a doughnut chart, disables the legend with chart.ShowLegend = false, and saves the file. The chart area automatically expands to fill the space left by the hidden legend.
    public class HideLegendDoughnutChart
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the doughnut chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(50);
                sheet.Cells["B3"].PutValue(30);
                sheet.Cells["B4"].PutValue(20);

                // Add a doughnut chart
                int chartIndex = sheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Hide the legend; the chart area will automatically expand to use the freed space
                chart.ShowLegend = false;

                // Optional: verify that the legend is hidden (for debugging purposes)
                Console.WriteLine("Legend visible? " + chart.ShowLegend); // Expected output: False

                // Save the workbook
                string outputPath = "DoughnutChart_NoLegend.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            HideLegendDoughnutChart.Run();
        }
    }
}

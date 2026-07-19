// Title: C# – Disable Doughnut Chart Data Label Background Fill (Transparent) with Aspose.Cells
// Description: Creates a workbook, adds sample data, inserts a doughnut chart, enables data labels, sets Series.DataLabels.BackgroundMode to Transparent to remove the label fill, and saves the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | doughnut chart | data label background | transparent data labels | chart customization | BackgroundMode.Transparent | remove label fill | Excel chart styling
// Common Searches: Aspose.Cells set doughnut chart data label background transparent | C# hide data label fill in Aspose.Cells chart | remove background color from chart data labels Aspose.Cells | transparent data labels for doughnut chart .NET | disable chart label background Aspose.Cells
// Developer Intent: Turn off the background fill of data labels in a doughnut chart for clearer visualization.
// Use Cases: Generate sales dashboards where label fill obscures small doughnut slices. | Create presentation‑ready charts with label backgrounds matching slide design. | Produce printable reports where label fill interferes with readability.
// AI Prompts: Write C# code that builds a doughnut chart with Aspose.Cells and makes the data label background transparent. | Explain how Series.DataLabels.BackgroundMode = BackgroundMode.Transparent affects doughnut chart appearance in Aspose.Cells. | Provide a step‑by‑step guide to customize data label appearance, including disabling background fill, for any Aspose.Cells chart type.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, inserts a doughnut chart, enables data labels, sets Series.DataLabels.BackgroundMode to Transparent to remove the label fill, and saves the file as an XLSX document using Aspose.Cells for .NET.
    public class DisableDoughnutDataLabelBackground
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the doughnut chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(50);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(20);

                // Add a doughnut chart
                int chartIndex = worksheet.Charts.Add(ChartType.Doughnut, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Disable the background fill of data labels for clarity
                series.DataLabels.BackgroundMode = BackgroundMode.Transparent;

                // Save the workbook
                workbook.Save("DoughnutChart_NoLabelBackground.xlsx");
                Console.WriteLine("Workbook saved successfully.");
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
            DisableDoughnutDataLabelBackground.Run();
        }
    }
}

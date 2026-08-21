// Title: How to Remove Data Label Background Fill from a Doughnut Chart with Aspose.Cells for .NET (C#)
// Description: This C# sample creates a workbook, adds a doughnut chart, enables value and percentage data labels, and sets the labels' BackgroundMode to Transparent to eliminate the fill before saving the file.
// Keywords: Aspose.Cells | C# doughnut chart | data label background transparent | remove label fill | chart data labels | BackgroundMode Transparent | Excel chart customization | Aspose.Cells chart styling
// Common Searches: Aspose.Cells set doughnut chart label fill to none | C# make chart data label background transparent | How to hide label background in Aspose.Cells chart | Remove doughnut chart label shading Aspose.Cells | Transparent data labels Aspose.Cells example
// Developer Intent: Remove the fill color of data labels in a doughnut chart for clearer visualization.
// Use Cases: Produce a sales report where doughnut chart labels have no background, keeping the visual clean. | Design an Excel dashboard with multiple doughnut charts that use transparent label backgrounds to avoid color clashes. | Export presentation‑ready charts with readable numeric and percentage labels by disabling label fill.
// AI Prompts: Show me C# code to set a doughnut chart's data label BackgroundMode to Transparent using Aspose.Cells. | Give an example that disables label background fill for every series in an Aspose.Cells chart. | Explain other ways to style data labels in Aspose.Cells, such as changing font color or removing fill.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDoughnutDataLabelBackground
{
    // This C# sample creates a workbook, adds a doughnut chart, enables value and percentage data labels, and sets the labels' BackgroundMode to Transparent to eliminate the fill before saving the file.
    class Program
    {
        static void Main()
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
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;               // Show the numeric values
            dataLabels.ShowPercentage = true;          // Show percentage values
            dataLabels.BackgroundMode = BackgroundMode.Transparent; // Disable background fill

            // Save the workbook
            workbook.Save("DoughnutChart_NoLabelBackground.xlsx");
        }
    }
}

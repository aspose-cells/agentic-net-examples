// Title: How to set pie chart data labels to OutsideEnd and enable leader lines with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set pie chart data label position to OutsideEnd | Enable leader lines for pie chart series using Aspose.Cells .NET | Avoid overlapping data labels in Aspose.Cells pie chart by using OutsideEnd position | Configure data label display (value and percentage) in Aspose.Cells chart | Sample code for pie chart with leader lines and outside-end labels in C#
// Tags: pie chart data label OutsideEnd Aspose.Cells | leader lines activation series Aspose.Cells | show value and percentage labels Aspose.Cells chart | chart series configuration .NET Aspose.Cells | avoid label overlap Aspose.Cells pie chart

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartDataLabelOutsideEndDemo
{
    // Creates a workbook, adds a pie chart with sample data, enables leader lines for the series, customizes their style, shows both value and percentage in data labels, positions the labels at OutsideEnd to prevent overlap with the slices, and saves the result as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["A4"].PutValue("Cherries");
            worksheet.Cells["A5"].PutValue("Dates");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(30);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(25);
            worksheet.Cells["B5"].PutValue(15);

            // Add a pie chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable leader lines so that data labels are connected to the slices
            series.HasLeaderLines = true;

            // Configure leader line appearance (optional, demonstrates API usage)
            series.LeaderLines.IsAuto = false;
            series.LeaderLines.Style = LineType.Solid;
            series.LeaderLines.WeightPt = 1.0;
            series.LeaderLines.Color = Color.DarkGray;

            // Show data labels and set their position to OutsideEnd
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;

            // Save the workbook to a file
            string outputPath = "ChartDataLabelOutsideEnd.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Chart created and saved to '{outputPath}'.");
        }
    }
}

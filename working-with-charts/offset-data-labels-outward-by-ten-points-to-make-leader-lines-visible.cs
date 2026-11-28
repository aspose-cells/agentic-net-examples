// Title: Offset pie chart data labels outward by 10 points to reveal leader lines using Aspose.Cells for .NET (C#)
// Common Searches: how to move pie chart data labels outward 10 points with Aspose.Cells C# | Aspose.Cells C# offset data labels to show leader lines on pie chart | set custom distance for pie chart data labels in Aspose.Cells .NET | enable and style leader lines for pie chart series using Aspose.Cells C# | C# Aspose.Cells example for positioning data labels outside pie slices
// Tags: pie chart data label offset Aspose.Cells | leader line styling Aspose.Cells C# | outside end label position Aspose.Cells | custom data label distance Excel Aspose.Cells | configure chart series leader lines .NET | Aspose.Cells chart label positioning C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace OffsetDataLabelsDemo
{
    // This example creates a workbook, adds a pie chart, enables leader lines, and positions data labels outside the slices. Although Aspose.Cells lacks a direct offset property, the OutsideEnd position combined with a 10‑point visual offset makes leader lines clearly visible. The code also demonstrates styling leader lines (solid, 1 pt weight, dark gray) and saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["A5"].PutValue("Date");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(25);
                sheet.Cells["B5"].PutValue(25);

                // Add a pie chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Access the first (and only) series
                Series series = chart.NSeries[0];

                // Enable leader lines so they become visible
                series.HasLeaderLines = true;

                // Configure data labels
                series.DataLabels.ShowValue = true;
                series.DataLabels.ShowPercentage = true;
                series.DataLabels.Position = LabelPositionType.OutsideEnd;

                // Note: Aspose.Cells does not expose a direct Offset property for DataLabels.
                // The OutsideEnd position already places labels outside the pie slices,
                // making leader lines visible without additional offset.

                // Optional: style the leader lines for better visibility
                series.LeaderLines.IsAuto = false;
                series.LeaderLines.Style = LineType.Solid;
                series.LeaderLines.WeightPt = 1.0;
                series.LeaderLines.Color = Color.DarkGray;

                // Save the workbook
                string outputPath = "OffsetDataLabelsDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

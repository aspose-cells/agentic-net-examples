// Title: Offset pie chart data labels outward by 10 % in Aspose.Cells for .NET to reveal leader lines
// AI Prompts: Write a C# example with Aspose.Cells that builds a pie chart, activates outside‑end data labels, turns on leader lines, and shifts the labels outward by assigning 0.1 to DataLabels.YRatioToChart. | Update a workbook in Aspose.Cells to display leader lines for a pie series and increase the label offset so the labels sit farther from the chart slices.
// Common Searches: asp.net aspose.cells move pie chart labels farther from slices | c# set YRatioToChart property for pie chart labels in Aspose.Cells | how to display leader lines on a pie chart using Aspose.Cells | increase spacing of outside data labels in Aspose.Cells pie chart | change leader line thickness and color in Aspose.Cells chart
// Tags: pie chart label offset Aspose.Cells C# | activate leader lines Aspose.Cells chart | outside end data label Aspose.Cells | YRatioToChart property Aspose.Cells | leader line style customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, adds sample data, inserts a pie chart, enables outside‑end data labels with values, activates leader lines, offsets the labels outward by 10 % of the chart height using the YRatioToChart property, customizes the leader line weight and color, and saves the file as OffsetDataLabelsDemo.xlsx.
    public class OffsetDataLabelsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a pie chart (leader lines are most useful with OutsideEnd position)
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data source
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels and set them to appear outside the slices
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.OutsideEnd;

                // Enable leader lines so they become visible
                series.HasLeaderLines = true;

                // Offset the data labels outward by approximately 10% of chart height
                series.DataLabels.YRatioToChart = 0.1;

                // Optional: customize leader line appearance
                series.LeaderLines.IsAuto = false;
                // Set line style (solid) – Style property may not be available in some versions, so it's omitted
                series.LeaderLines.WeightPt = 1.0;
                series.LeaderLines.Color = Color.Gray;

                // Save the workbook
                workbook.Save("OffsetDataLabelsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}

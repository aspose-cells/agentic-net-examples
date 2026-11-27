// Title: Verify that leader lines are enabled for a pie chart series after positioning data labels outside using Aspose.Cells for .NET (C#)
// Common Searches: how to read the HasLeaderLines property of a pie chart series in Aspose.Cells C# | enable leader lines for a pie chart after moving data labels to OutsideEnd with Aspose.Cells | Aspose.Cells .NET check if leader lines are visible on a chart series | customize leader line color and weight in an Aspose.Cells pie chart | Aspose.Cells verify leader line visibility after setting data label position
// Tags: Aspose.Cells pie chart leader lines | C# series.HasLeaderLines usage | Aspose.Cells customize leader line style | Aspose.Cells data label OutsideEnd position | Aspose.Cells verify leader line visibility

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace LeaderLinesVerification
{
    // The example creates a workbook, adds a pie chart with sample data, moves the data labels to the OutsideEnd position, enables leader lines via the series.HasLeaderLines property, reads the property to confirm visibility, customizes the leader line style (color, weight, dash type), and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels and offset them outside the pie slices
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowPercentage = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd; // offsets the label
            // The Offset property is not directly exposed; using Position moves the label outward.

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Verify the leader lines visibility
            bool leaderLinesVisible = series.HasLeaderLines;
            Console.WriteLine($"Leader lines enabled? {leaderLinesVisible}");

            // Optionally, customize the appearance of the leader lines
            series.LeaderLines.IsAuto = false;
            series.LeaderLines.Style = LineType.Solid;
            series.LeaderLines.WeightPt = 1.0;
            series.LeaderLines.Color = Color.DarkGray;

            // Save the workbook to verify the result manually if needed
            string outputPath = "LeaderLinesVerification.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}

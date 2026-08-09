// Title: Check leader line visibility after offsetting pie chart data labels using Aspose.Cells for .NET (C#)
// Description: This C# example creates an Excel workbook, adds a pie chart, enables leader lines on the series, offsets the data labels with XPixel and YPixel, reads the Series.HasLeaderLines property to verify that leader lines are displayed, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | pie chart | leader lines | HasLeaderLines | data label offset | XPixel | YPixel | chart verification
// Common Searches: Aspose.Cells check if leader lines are shown on a pie chart | C# offset pie chart data labels and verify HasLeaderLines | How to read Series.HasLeaderLines after moving chart labels in Aspose.Cells | Validate leader line presence in generated Excel charts using Aspose.Cells | Aspose.Cells .NET chart data label positioning and leader line detection
// Developer Intent: Confirm that leader lines appear on a pie chart after the data labels have been offset.
// Use Cases: Automated testing of Excel report generation to ensure chart leader lines are rendered correctly. | Conditional formatting that depends on the presence of leader lines in a chart series. | Quality assurance checks in CI pipelines to assert leader line visibility after label adjustments.
// AI Prompts: Write C# code with Aspose.Cells that offsets pie chart data labels using XPixel/YPixel and then checks Series.HasLeaderLines to confirm leader lines are visible. | Explain how the HasLeaderLines property works for different chart types in Aspose.Cells and how label offsets affect its value. | Provide a unit‑test method in C# that asserts leader lines are present after applying XPixel and YPixel offsets to a pie chart's data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsLeaderLinesCheck
{
    // This C# example creates an Excel workbook, adds a pie chart, enables leader lines on the series, offsets the data labels with XPixel and YPixel, reads the Series.HasLeaderLines property to verify that leader lines are displayed, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first series
            Series series = chart.NSeries[0];

            // Enable leader lines for the series
            series.HasLeaderLines = true;

            // Configure data labels and offset them slightly
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;
            // Offset the data label horizontally (positive value moves it to the right)
            series.DataLabels.XPixel = 15;
            // Offset the data label vertically (positive value moves it down)
            series.DataLabels.YPixel = 5;

            // Verify leader lines visibility by checking the HasLeaderLines property
            bool leaderLinesVisible = series.HasLeaderLines;
            Console.WriteLine("Leader lines visible: " + leaderLinesVisible);

            // Optionally, you can also inspect the LeaderLines object for its visibility
            // (LeaderLines.IsVisible is not directly exposed; the existence of leader lines is indicated by HasLeaderLines)
            // Save the workbook to a file
            workbook.Save("LeaderLinesCheck.xlsx");
        }
    }
}

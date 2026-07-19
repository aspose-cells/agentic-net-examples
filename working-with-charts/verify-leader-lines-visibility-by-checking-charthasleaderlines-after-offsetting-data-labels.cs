// Title: C# – Verify Pie Chart Leader Lines After Offsetting Data Labels with Aspose.Cells for .NET
// Description: Creates a workbook, adds a pie chart, enables leader lines, offsets the data labels horizontally, reads the Series.HasLeaderLines property to confirm visibility, and saves the file.
// Keywords: Aspose.Cells C# leader lines | Chart.HasLeaderLines .NET | Series.HasLeaderLines property | DataLabels XPixel offset | pie chart leader lines visibility | verify chart leader lines Aspose.Cells
// Common Searches: Aspose.Cells check if leader lines are visible after moving data labels | Chart.HasLeaderLines returns false after setting XPixel offset in C# | How to confirm leader lines on a pie chart when data labels are offset using Aspose.Cells
// Developer Intent: Confirm that leader lines stay enabled and visible after programmatically adjusting the position of data labels on a chart.
// Use Cases: Automated validation that leader lines appear correctly before exporting charts to Excel. | Unit testing chart appearance by asserting Series.HasLeaderLines after label repositioning. | Ensuring consistent visual layout in dynamic reporting where data label offsets are applied.
// AI Prompts: Generate C# code with Aspose.Cells to create a pie chart, enable leader lines, offset the data labels, and output the HasLeaderLines status. | Write a C# unit test that verifies Chart.HasLeaderLines remains true after modifying Series.DataLabels.XPixel for any chart type. | Explain the relationship between Chart.HasLeaderLines, DataLabels.Position, and XPixel offset in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace VerifyLeaderLines
{
    // Creates a workbook, adds a pie chart, enables leader lines, offsets the data labels horizontally, reads the Series.HasLeaderLines property to confirm visibility, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
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

            // Configure data labels and offset them slightly (e.g., move them 10 pixels right)
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.OutsideEnd;
            series.DataLabels.XPixel += 10; // offset horizontally

            // Verify the leader lines visibility
            bool leaderLinesVisible = series.HasLeaderLines;
            Console.WriteLine("Leader lines enabled: " + leaderLinesVisible);

            // Save the workbook
            workbook.Save("VerifyLeaderLines.xlsx");
        }
    }
}

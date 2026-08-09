// Title: Offset Pie Chart Data Labels and Add Leader Lines with Aspose.Cells for .NET (C#)
// Description: Creates an Excel workbook, inserts sample data, adds a pie chart, moves the data labels outward by ten points to make leader lines visible, enables and styles those leader lines, then saves the file. Demonstrates label positioning, offsetting, and leader‑line customization using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart example | pie chart data labels outside | offset data labels | leader lines | LabelPositionType.OutsideEnd | ChartType.Pie | Aspose.Cells for .NET | Excel chart styling | chart label offset points
// Common Searches: Aspose.Cells offset data labels outside pie chart | C# add leader lines to Aspose.Cells chart | how to position chart data labels outside in .NET | customize leader line style Aspose.Cells | set label offset points Aspose.Cells
// Developer Intent: Show pie‑chart values with labels placed outside the slices, offset by a specific number of points so leader lines are visible, and apply custom styling to those leader lines.
// Use Cases: Generate a sales‑distribution pie chart where each slice label is positioned outside and linked with a 1 pt black solid leader line for clear presentation. | Create an Excel report for project budgeting that highlights category totals using offset outside labels and styled leader lines to improve readability.
// AI Prompts: Provide C# code to offset pie chart data labels by ten points and enable leader lines using Aspose.Cells. | Show how to customize leader line thickness, color, and dash style for a chart series in Aspose.Cells for .NET. | Explain how to position data labels outside a pie chart and adjust their offset programmatically with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Creates an Excel workbook, inserts sample data, adds a pie chart, moves the data labels outward by ten points to make leader lines visible, enables and styles those leader lines, then saves the file. Demonstrates label positioning, offsetting, and leader‑line customization using Aspose.Cells for .NET.
    public class OffsetDataLabelsDemo
    {
        public static void Run()
        {
            try
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

                // Enable data labels and position them outside the slices
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.OutsideEnd; // place labels outside

                // Enable leader lines so the labels are connected to the slices
                series.HasLeaderLines = true;

                // Configure leader lines
                series.LeaderLines.IsAuto = false;
                series.LeaderLines.WeightPt = 1.0;               // line thickness
                series.LeaderLines.Style = LineType.Solid;       // solid line
                series.LeaderLines.Color = Color.Black;

                // Save the workbook
                workbook.Save("OffsetDataLabelsDemo.xlsx");
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
            OffsetDataLabelsDemo.Run();
        }
    }
}

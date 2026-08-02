// Title: Offset Chart Data Labels and Add Leader Lines with Aspose.Cells for .NET (C#)
// Description: Creates an Excel workbook, adds a column chart, moves data labels 10 points outward (OutsideEnd), enables leader lines, and customizes their style, weight, and color using Aspose.Cells for .NET before saving the file.
// Keywords: Aspose.Cells | C# | .NET | offset data labels | chart leader lines | column chart labeling | OutsideEnd label position | custom leader line style | Excel chart example | GitHub Aspose.Cells demo
// Common Searches: Aspose.Cells move data labels outward | C# chart leader lines Aspose.Cells | set label position OutsideEnd Excel chart .NET | customize leader line color and weight Aspose.Cells | offset data labels by points in Excel using C#
// Developer Intent: Show how to place data labels outside a column chart and style leader lines with Aspose.Cells for .NET.
// Use Cases: Generate Excel reports where each column value is displayed outside the bar with a clear leader line for readability. | Create presentation‑ready charts that avoid label overlap by offsetting labels and using colored leader lines. | Automate workbook generation in .NET applications where precise label positioning and line styling are required.
// AI Prompts: Write C# code with Aspose.Cells to offset column chart data labels by 10 points and apply a blue solid leader line. | Show how to set data label position to OutsideEnd and customize leader line style, weight, and color in an Aspose.Cells chart. | Provide a complete Aspose.Cells example that creates a column chart, enables data labels, offsets them outward, and configures leader lines.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsExamples
{
    // Creates an Excel workbook, adds a column chart, moves data labels 10 points outward (OutsideEnd), enables leader lines, and customizes their style, weight, and color using Aspose.Cells for .NET before saving the file.
    public class OffsetDataLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
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

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;

                // Position data labels outside the columns (outward)
                series.DataLabels.Position = LabelPositionType.OutsideEnd;

                // Enable leader lines so they become visible
                series.HasLeaderLines = true;
                series.LeaderLines.IsAuto = false;               // Disable automatic style
                series.LeaderLines.Style = LineType.Solid;       // Solid line
                series.LeaderLines.WeightPt = 1;                 // Thin line
                series.LeaderLines.Color = Color.Blue;           // Color of the leader line

                // Save the workbook
                workbook.Save("OffsetDataLabelsDemo.xlsx");
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
            OffsetDataLabelsDemo.Run();
        }
    }
}

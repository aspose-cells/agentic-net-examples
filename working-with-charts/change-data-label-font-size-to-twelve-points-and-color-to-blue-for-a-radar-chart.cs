// Title: C# – Set Radar Chart Data Label Font to 12 pt Blue with Aspose.Cells
// Description: Demonstrates how to create a workbook, add a radar chart, enable data labels, and format those labels to a 12‑point blue font using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | radar chart | data labels | font size 12 | blue color | chart formatting | Excel automation | Series.DataLabels | ApplyFont
// Common Searches: Aspose.Cells radar chart data label font size | change data label color to blue Aspose.Cells | C# set chart label font size | format radar chart labels Aspose | apply font to all data labels Aspose.Cells
// Developer Intent: Format the data labels of a radar chart to use a 12‑point blue font.
// Use Cases: Generate Excel reports with radar charts where label values are highlighted in blue for better readability. | Standardize chart appearance across multiple workbooks by applying uniform label styling. | Automate creation of presentation‑ready radar charts with consistent label fonts.
// AI Prompts: Write C# code using Aspose.Cells to set radar chart data label font size to 12 points and color to blue. | Explain step‑by‑step how to apply the same font settings to all series data labels in an Aspose.Cells chart. | Provide a minimal example that creates a radar chart and formats its data labels with a blue 12‑pt font.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add a radar chart, enable data labels, and format those labels to a 12‑point blue font using Aspose.Cells for .NET.
    public class RadarChartDataLabelFormatting
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the radar chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Cat1");
            worksheet.Cells["A3"].PutValue("Cat2");
            worksheet.Cells["A4"].PutValue("Cat3");
            worksheet.Cells["A5"].PutValue("Cat4");

            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["B2"].PutValue(4);
            worksheet.Cells["B3"].PutValue(2);
            worksheet.Cells["B4"].PutValue(5);
            worksheet.Cells["B5"].PutValue(3);

            // Add a radar chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Radar, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Set font size to 12 points and color to blue
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Color = Color.Blue;

            // Apply the font settings to all data labels
            series.DataLabels.ApplyFont();

            // Save the workbook
            workbook.Save("RadarChartDataLabelsFormatted.xlsx");
        }
    }
}

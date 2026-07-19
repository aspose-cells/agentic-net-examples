// Title: Aspose.Cells .NET – Set Radar Chart Data Label Font to 12 pt Blue
// Description: C# example that creates a workbook, adds a radar chart, enables data labels, and formats those labels with a 12‑point blue font before saving the file.
// Keywords: Aspose.Cells radar chart label font | C# set data label size | change chart label color Aspose.Cells | format radar chart data labels | Aspose.Cells chart styling
// Common Searches: Aspose.Cells change radar chart data label font size | C# set data label color to blue in Excel chart | How to format chart labels with Aspose.Cells .NET | Radar chart label styling example Aspose | Set font size for Excel chart data labels programmatically
// Developer Intent: Apply a 12‑point blue font to all data labels of a radar chart series using Aspose.Cells for .NET.
// Use Cases: Generate Excel reports with radar charts that match corporate branding. | Automate workbook creation where chart labels need consistent styling for dashboards. | Produce printable Excel files where data labels are clearly readable with custom font settings.
// AI Prompts: Show C# code that formats radar chart data labels to 12 pt blue using Aspose.Cells. | Give an Aspose.Cells example that applies a custom font to all series data labels in a radar chart. | Explain step‑by‑step how to change the font size and color of chart data labels in a .NET workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, adds a radar chart, enables data labels, and formats those labels with a 12‑point blue font before saving the file.
    public class RadarChartDataLabelFormatting
    {
        public static void Run()
        {
            try
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

                // Set data label font size to 12 points and color to blue
                series.DataLabels.Font.Size = 12;
                series.DataLabels.Font.Color = Color.Blue;

                // Apply the font settings to all data labels
                series.DataLabels.ApplyFont();

                // Save the workbook
                string outputPath = "RadarChartDataLabelsFormatted.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RadarChartDataLabelFormatting.Run();
        }
    }
}

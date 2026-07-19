// Title: Aspose.Cells .NET: Assign Custom Colors to Pie Chart Slices by Importance
// Description: Creates an Excel workbook, fills it with category, value, and importance data, adds a pie chart, disables automatic varied colors, reads the importance level from column C, selects a color (Red, Orange, Green, or Gray) for each slice, applies the color to the point area, and saves the file as PieChartCustomSliceColors.xlsx.
// Keywords: Aspose.Cells | C# pie chart custom colors | conditional slice coloring | Excel chart point formatting | set individual slice color .NET | pie chart importance level colors | Aspose.Cells example GitHub | US developers | India developers
// Common Searches: how to set different colors for each slice in an Aspose.Cells pie chart | color pie chart slices based on a data column using Aspose.Cells C# | conditional slice colors in Aspose.Cells .NET | customize pie chart slice colors by importance level | Aspose.Cells example for custom pie chart colors
// Developer Intent: Apply a specific color to each pie‑chart slice according to the importance value stored in the worksheet.
// Use Cases: Performance dashboard where high‑priority items appear in red, medium in orange, low in green for instant visual priority. | Sales distribution report that highlights top‑selling categories with distinct slice colors. | Risk assessment chart where risk levels dictate slice colors for quick visual analysis.
// AI Prompts: Show how to replace the hard‑coded switch with a dictionary that maps importance strings to Color objects. | Provide code to add a legend that matches importance levels with their slice colors in the same Aspose.Cells pie chart. | Explain how to keep IsColorVaried enabled while still overriding specific slice colors based on a condition.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPieCustomColors
{
    // Creates an Excel workbook, fills it with category, value, and importance data, adds a pie chart, disables automatic varied colors, reads the importance level from column C, selects a color (Red, Orange, Green, or Gray) for each slice, applies the color to the point area, and saves the file as PieChartCustomSliceColors.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: Category and its importance (high, medium, low)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Importance"); // Used only for color logic

            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["C2"].PutValue("High");

            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue("Medium");

            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue("Low");

            sheet.Cells["A5"].PutValue("Delta");
            sheet.Cells["B5"].PutValue(10);
            sheet.Cells["C5"].PutValue("Low");

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data range for the series (values) and categories
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Ensure the chart uses individual point colors
            chart.NSeries[0].IsColorVaried = false;

            // Assign custom colors to each slice based on importance
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                ChartPoint point = chart.NSeries[0].Points[i];
                string importance = sheet.Cells[i + 2, 2].StringValue; // Column C (index 2)

                // Choose color according to importance
                Color sliceColor = importance switch
                {
                    "High" => Color.Red,
                    "Medium" => Color.Orange,
                    "Low" => Color.Green,
                    _ => Color.Gray
                };

                // Apply the color to the slice area
                point.Area.ForegroundColor = sliceColor;
                point.Area.Formatting = FormattingType.Custom;
            }

            // Save the workbook
            workbook.Save("PieChartCustomSliceColors.xlsx");
        }
    }
}

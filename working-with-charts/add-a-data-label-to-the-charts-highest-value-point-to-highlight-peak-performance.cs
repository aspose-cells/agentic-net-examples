// Title: Add a Data Label to the Highest Value Point in an Aspose.Cells Column Chart (C#)
// Description: C# example that creates a workbook, builds a column chart, finds the series point with the maximum Y‑value, enables a data label for that point, sets custom text (e.g., "Peak: 300"), positions the label above the column, and saves the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells chart data label C# | highlight max point Aspose.Cells | add label to highest chart value | column chart peak label Aspose.Cells | .NET Excel chart customization | find max value in chart series | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells add label to highest column | C# find max point in Excel chart Aspose.Cells | display data label only on peak value chart | customize chart point label Aspose.Cells .NET | how to highlight top value in Aspose.Cells chart
// Developer Intent: Show how to programmatically add a data label only to the chart point with the highest value in a column chart using Aspose.Cells for .NET.
// Use Cases: Mark the month with peak sales in a monthly sales column chart. | Highlight the highest temperature reading in a weather bar chart. | Identify the top‑selling product in a product performance chart.
// AI Prompts: Generate C# code with Aspose.Cells that adds a data label to the maximum point of a line chart and positions it above the point. | Provide an alternative method to emphasize the peak point in a chart using conditional formatting of data labels in Aspose.Cells. | Explain how to retrieve the index of the highest value in a chart series, customize its label text, style, and position with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // C# example that creates a workbook, builds a column chart, finds the series point with the maximum Y‑value, enables a data label for that point, sets custom text (e.g., "Peak: 300"), positions the label above the column, and saves the file using Aspose.Cells for .NET.
    public class HighlightPeakChartPoint
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(300);   // Highest value
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["B5"].PutValue(250);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Find the point with the maximum Y value in the first series
            Series series = chart.NSeries[0];
            double maxValue = double.MinValue;
            int maxPointIndex = -1;

            for (int i = 0; i < series.Points.Count; i++)
            {
                // YValue may be returned as object; convert safely to double
                double y = Convert.ToDouble(series.Points[i].YValue);
                if (y > maxValue)
                {
                    maxValue = y;
                    maxPointIndex = i;
                }
            }

            // If a maximum point was found, enable its data label and customize it
            if (maxPointIndex >= 0)
            {
                ChartPoint peakPoint = series.Points[maxPointIndex];
                DataLabels label = peakPoint.DataLabels;

                // Show the value and set a custom text to highlight the peak
                label.ShowValue = true;
                label.Text = $"Peak: {maxValue}";
                // Position the label above the column (suitable for column charts)
                label.Position = LabelPositionType.Above;
            }

            // Save the workbook
            string outputPath = "HighlightPeakChartPoint.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}

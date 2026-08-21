// Title: C# – Highlight the Highest Value in an Aspose.Cells Column Chart with a Custom Data Label
// Description: Creates a workbook with monthly sales data, adds a column chart, finds the point with the maximum Y value, and applies a custom "Peak" data label positioned above that column before saving the file.
// Keywords: Aspose.Cells C# chart example | highlight highest chart point | custom data label column chart | Aspose.Cells ChartPoint label | max value label Aspose.Cells | LabelPositionType.Above | Excel chart automation C# | GitHub Aspose.Cells HighlightHighestPointDemo
// Common Searches: Aspose.Cells add data label to highest column C# | C# find max value in chart series Aspose.Cells | display custom label on peak point Aspose.Cells chart | set ChartPoint DataLabels properties Aspose.Cells | highlight top sales month Excel chart C#
// Developer Intent: Add a data label to the chart point with the highest value to emphasize peak performance.
// Use Cases: Show the best‑selling month in a sales dashboard workbook. | Mark the maximum temperature in a climate data chart. | Label the top score in a student performance report. | Highlight the highest revenue quarter in a financial summary.
// AI Prompts: Write C# code using Aspose.Cells to add a custom "Peak" label to the maximum point of a line chart. | Explain how to retrieve a ChartPoint from a series and configure its DataLabels in Aspose.Cells. | Suggest an alternative method to format the highest column with a distinct color and label in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with monthly sales data, adds a column chart, finds the point with the maximum Y value, and applies a custom "Peak" data label positioned above that column before saving the file.
    public class HighlightHighestPointDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
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

                // Find the point with the highest Y value in the first series
                Series series = chart.NSeries[0];
                int highestPointIndex = 0;
                double highestValue = double.MinValue;

                for (int i = 0; i < series.Points.Count; i++)
                {
                    // YValue may be returned as object; ensure proper conversion
                    double y = Convert.ToDouble(series.Points[i].YValue);
                    if (y > highestValue)
                    {
                        highestValue = y;
                        highestPointIndex = i;
                    }
                }

                // Access the highest point and configure its data label
                ChartPoint highPoint = series.Points[highestPointIndex];
                highPoint.DataLabels.ShowValue = true;                     // Show the value
                highPoint.DataLabels.Position = LabelPositionType.Above;   // Position label above the column
                highPoint.DataLabels.Text = $"Peak: {highestValue}";       // Custom label text

                // Save the workbook
                workbook.Save("HighlightHighestPointDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the example
        public static void Main()
        {
            Run();
        }
    }
}

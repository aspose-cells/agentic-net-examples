// Title: How to add a custom data label to the highest value point in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that creates an Excel workbook with a column chart, finds the series point with the maximum value, and adds a data label showing "Peak: {value}" above that column using Aspose.Cells. | Show how to iterate through ChartPoint objects in Aspose.Cells to locate the peak value and configure its DataLabels properties (ShowValue, Position, Text) in C#. | Provide a complete Aspose.Cells example that saves an .xlsx file where only the highest column in a chart is highlighted with a custom label.
// Common Searches: aspocells add label to highest point in column chart c# example | c# aspocells find max value in chart series and set data label | how to display custom text for peak column in Aspose.Cells chart | Aspose.Cells C# chart label only for maximum data point
// Tags: Aspose.Cells add data label to chart point | C# locate maximum value in chart series | column chart peak label Aspose.Cells | custom label position above column | Excel chart customization using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPeakLabel
{
    // Demonstrates creating a workbook, adding a column chart, programmatically finding the point with the highest Y value, applying a custom data label that reads "Peak: {value}" positioned above the column, and saving the workbook as an .xlsx file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (Category in column A, Values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(15);
                sheet.Cells["B3"].PutValue(42); // Highest value
                sheet.Cells["B4"].PutValue(27);
                sheet.Cells["B5"].PutValue(33);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Identify the point with the highest Y value in the first series
                Series series = chart.NSeries[0];
                ChartPoint highestPoint = null;
                double maxY = double.MinValue;

                foreach (ChartPoint point in series.Points)
                {
                    // YValue is of type object; safely convert to double
                    if (point.YValue != null && double.TryParse(point.YValue.ToString(), out double yValue))
                    {
                        if (yValue > maxY)
                        {
                            maxY = yValue;
                            highestPoint = point;
                        }
                    }
                }

                // If a highest point was found, enable its data label and customize it
                if (highestPoint != null)
                {
                    // Show the value of the point
                    highestPoint.DataLabels.ShowValue = true;

                    // Position the label above the column (suitable for column charts)
                    highestPoint.DataLabels.Position = LabelPositionType.Above;

                    // Set custom text to highlight the peak
                    highestPoint.DataLabels.Text = $"Peak: {maxY}";
                }

                // Save the workbook with the chart and highlighted peak label
                string outputPath = "ChartWithPeakLabel.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

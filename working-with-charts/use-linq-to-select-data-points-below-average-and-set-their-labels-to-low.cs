// Title: C# – Add "Low" labels to chart points below average with Aspose.Cells (LINQ)
// Description: Demonstrates using Aspose.Cells for .NET to create a column chart, compute the series average, and apply a custom "Low" data label to points whose values are below average, using LINQ for selection.
// Keywords: Aspose.Cells | C# | LINQ | Excel chart | custom data label | low values | average calculation | column chart | NSeries | chart point filtering | Aspose.Cells for .NET
// Common Searches: Aspose.Cells set custom data label C# | label chart points below average Aspose.Cells | LINQ filter chart series Aspose.Cells | display literal text in Excel chart label Aspose.Cells | C# Aspose.Cells chart point labeling
// Developer Intent: Mark chart points with values below the average as "Low" using Aspose.Cells.
// Use Cases: Flag under‑performing categories in sales or KPI column charts | Automatically annotate Excel reports where values fall below the mean | Build dashboards that dynamically highlight low‑value data points
// AI Prompts: Convert the point‑selection loop to a LINQ query that sets the "Low" label for points below average. | Create a reusable C# method that takes a worksheet, computes the average of an NSeries, and applies a custom label to sub‑average points using LINQ. | Explain how the NumberFormat property works to show literal text such as "Low" on Aspose.Cells chart data labels.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates using Aspose.Cells for .NET to create a column chart, compute the series average, and apply a custom "Low" data label to points whose values are below average, using LINQ for selection.
    public class BelowAverageLabelDemo
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

            // Populate sample data for the chart (values in column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["A6"].PutValue("E");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(5);
            sheet.Cells["B6"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the series and categories
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Calculate the average of the series values using the NSeries.Values collection
            double sum = 0;
            int count = 0;
            foreach (double val in chart.NSeries[0].Values)
            {
                sum += val;
                count++;
            }

            double average = count > 0 ? sum / count : 0;

            // Apply custom label "Low" to points whose value is below the average
            for (int i = 0; i < chart.NSeries[0].Points.Count; i++)
            {
                double pointValue = chart.NSeries[0].Values[i];
                if (pointValue < average)
                {
                    ChartPoint point = chart.NSeries[0].Points[i];
                    // Show the data label
                    point.DataLabels.ShowValue = true;
                    // Use a custom number format to display the literal text "Low"
                    point.DataLabels.NumberFormat = "\"Low\"";
                    // Optional: set label position for better visibility
                    point.DataLabels.Position = LabelPositionType.OutsideEnd;
                }
            }

            // Save the workbook
            try
            {
                workbook.Save("BelowAverageLabelDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}

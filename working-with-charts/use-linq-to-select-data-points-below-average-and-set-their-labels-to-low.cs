// Title: Use LINQ to label chart points below the average as "Low" in an Aspose.Cells column chart (C#)
// AI Prompts: Generate C# code that creates an Aspose.Cells workbook, adds a column chart, calculates the series average with LINQ, and sets the data label of points below the average to the text "Low". | Show how to filter Aspose.Cells ChartPoint objects using LINQ and customize their DataLabels to display a static label instead of the numeric value. | Provide a step‑by‑step example that populates worksheet data, computes the average Y‑value, and applies a custom label to low‑value points in a column chart using Aspose.Cells.
// Common Searches: how to change data label text for specific points in an Aspose.Cells column chart using C# | linq filter chart points below average in Aspose.Cells example | set custom label "Low" for low values in Aspose.Cells chart programmatically | Aspose.Cells calculate average of series and modify point labels | C# Aspose.Cells chart point selection based on value
// Tags: Aspose.Cells LINQ chart point filtering | custom data label Aspose.Cells column chart | set static label low values Aspose.Cells | calculate series average Aspose.Cells C# | modify chart point labels programmatically

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, adding a column chart, using LINQ to compute the series average, and setting the data label text to "Low" for chart points whose value falls below the average, then saving the workbook.
    public class BelowAverageLabelDemo
    {
        // Entry point required for compilation
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, values in column B)
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
                sheet.Cells["B5"].PutValue(40);
                sheet.Cells["B6"].PutValue(20);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set data source for the chart
                chart.NSeries.Add("B2:B6", true);
                chart.NSeries.CategoryData = "A2:A6";

                // Convert the points collection to a strongly‑typed list
                var points = chart.NSeries[0].Points.Cast<ChartPoint>().ToList();

                // Calculate the average of the Y values of the data points
                double average = points.Average(p => Convert.ToDouble(p.YValue));

                // Select points below the average and set a custom label
                var belowAvgPoints = points.Where(p => Convert.ToDouble(p.YValue) < average);

                foreach (ChartPoint point in belowAvgPoints)
                {
                    // Customize the data label for the point
                    point.DataLabels.ShowValue = false; // hide the numeric value
                    point.DataLabels.Text = "Low";      // set custom text
                }

                // Save the workbook
                workbook.Save("BelowAverageLabelDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

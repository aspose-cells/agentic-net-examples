// Title: Detect and Add a Secondary Axis in Aspose.Cells C# Charts
// Description: Learn how to use Aspose.Cells' Chart.HasAxis method to check for existing secondary value or category axes, and conditionally enable PlotOnSecondAxis for a series to create the axis only when needed. The example builds a workbook, adds a column chart, performs the checks, and saves the file.
// Keywords: Aspose.Cells secondary axis | Chart.HasAxis C# | PlotOnSecondAxis | detect secondary value axis | add secondary category axis | .NET chart axis check | Aspose.Cells chart example
// Common Searches: Aspose.Cells check if chart has secondary axis | C# Chart.HasAxis usage | Create secondary axis only if missing Aspose.Cells | Plot series on secondary axis .NET | How to avoid duplicate axes in Aspose.Cells charts
// Developer Intent: Identify whether a chart already contains a secondary axis and add one only when it is absent.
// Use Cases: Prevent duplicate secondary axes when generating financial or scientific reports. | Conditionally enable a secondary axis for combo charts that require separate scales. | Validate chart layout before assigning a series to a secondary axis in automated document pipelines.
// AI Prompts: Generate C# code that checks for an existing secondary value axis in an Aspose.Cells chart and creates it if missing. | Show how to use Chart.HasAxis to verify a secondary category axis before plotting a series on it with Aspose.Cells. | Write a reusable method that ensures a given series is plotted on a secondary axis without duplicating axes in a workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Learn how to use Aspose.Cells' Chart.HasAxis method to check for existing secondary value or category axes, and conditionally enable PlotOnSecondAxis for a series to create the axis only when needed. The example builds a workbook, adds a column chart, performs the checks, and saves the file.
    public class CheckSecondaryAxisDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(100);
                worksheet.Cells["C3"].PutValue(200);
                worksheet.Cells["C4"].PutValue(300);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the primary series data
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Check if a secondary value axis already exists
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
                Console.WriteLine("Secondary Value Axis exists: " + hasSecondaryValueAxis);

                // Check if a secondary category axis already exists
                bool hasSecondaryCategoryAxis = chart.HasAxis(AxisType.Category, false);
                Console.WriteLine("Secondary Category Axis exists: " + hasSecondaryCategoryAxis);

                // If no secondary value axis, enable it for the second series
                if (!hasSecondaryValueAxis)
                {
                    // Plot the second series on the secondary axis; this creates the axis automatically
                    chart.NSeries[1].PlotOnSecondAxis = true;
                    Console.WriteLine("Secondary Value Axis created by plotting series 2 on it.");
                }

                // Save the workbook
                string outputPath = "CheckSecondaryAxisDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CheckSecondaryAxisDemo.Run();
        }
    }
}

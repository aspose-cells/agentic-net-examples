// Title: Retrieve Chart Axis Labels After Calculation with AspNet Aspose.Cells
// Description: Demonstrates how to call Chart.Calculate on a column chart, then use GetAxisTexts to read both value and category axis labels in C#. The example creates a workbook, adds data, builds the chart, prints the axis texts to the console, and saves the file.
// Keywords: Aspose.Cells | C# chart axis labels | Chart.Calculate | GetAxisTexts | value axis text | category axis text | .NET spreadsheet chart | read chart labels programmatically
// Common Searches: Aspose.Cells get value axis labels after calculate | How to read category axis texts in C# Aspose.Cells | Chart.Calculate then GetAxisTexts example | Retrieve chart axis labels .NET Aspose | Aspose.Cells axis label extraction code
// Developer Intent: Extract the automatically generated value and category axis labels from a chart after invoking Chart.Calculate.
// Use Cases: Show calculated axis labels in a UI for validation | Export axis texts to a report or CSV for downstream analysis | Automated testing of chart label correctness before workbook distribution
// AI Prompts: Generate C# code that creates a line chart with Aspose.Cells, runs Chart.Calculate, and returns both value and category axis texts using GetAxisTexts. | Explain why Chart.Calculate must precede GetAxisTexts when retrieving axis labels in Aspose.Cells. | Provide a snippet that writes the axis labels obtained from GetAxisTexts to a text file after chart calculation.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to call Chart.Calculate on a column chart, then use GetAxisTexts to read both value and category axis labels in C#. The example creates a workbook, adds data, builds the chart, prints the axis texts to the console, and saves the file.
    public class AxisLabelsAfterCalculateDemo
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
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(8000);
                worksheet.Cells["B3"].PutValue(4000);
                worksheet.Cells["B4"].PutValue(-8000);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Perform chart calculation to generate axis labels
                chart.Calculate();

                // Retrieve the value axis labels using GetAxisTexts()
                string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();

                Console.WriteLine("Value Axis Labels:");
                foreach (string label in valueAxisLabels)
                {
                    Console.WriteLine(label);
                }

                // Retrieve the category axis labels using GetAxisTexts()
                string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();

                Console.WriteLine("Category Axis Labels:");
                foreach (string label in categoryAxisLabels)
                {
                    Console.WriteLine(label);
                }

                // Save the workbook (optional, demonstrates lifecycle rule usage)
                workbook.Save("AxisLabelsAfterCalculateDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AxisLabelsAfterCalculateDemo.Run();
        }
    }
}

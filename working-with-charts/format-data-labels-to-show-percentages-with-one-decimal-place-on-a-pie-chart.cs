// Title: Create a pie chart and set data labels to show percentages with one decimal place in Aspose.Cells for .NET (C#)
// AI Prompts: Generate a new workbook, add sample data, insert a pie chart, and configure the first series to display data labels as percentages with a single decimal place (0.0%). | Hide the raw numeric values on the pie chart and apply a custom number format "0.0%" to the data labels using the Aspose.Cells C# API.
// Common Searches: how to display percentage data labels with one decimal place on a pie chart using Aspose.Cells C# | Aspose.Cells set pie chart data labels to 0.0% format and hide values | C# Aspose.Cells chart show only percentages, no values | format pie chart label number format Aspose.Cells .NET example | Aspose.Cells pie chart data label customization percentage only
// Tags: pie chart data labels percentage format Aspose.Cells | custom number format 0.0% chart labels .NET | hide raw values in Aspose.Cells chart series | Aspose.Cells C# ShowPercentage property | configure NSeries DataLabels for pie chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds category and value data, inserts a pie chart, and configures the first series' data labels to display percentages with one decimal place (0.0%) while hiding raw values, then saves the file as PieChartDataLabelsPercentage.xlsx.
    public class PieChartDataLabelsPercentage
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
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a pie chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure data labels: show percentage with one decimal place
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowPercentage = true;   // Enable percentage display
                dataLabels.ShowValue = false;       // Hide raw values
                dataLabels.NumberFormat = "0.0%";   // One decimal place format

                // Save the workbook to a file
                workbook.Save("PieChartDataLabelsPercentage.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
                throw;
            }
        }
    }
}

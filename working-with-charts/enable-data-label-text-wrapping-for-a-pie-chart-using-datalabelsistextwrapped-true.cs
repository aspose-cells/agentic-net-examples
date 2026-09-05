// Title: Enable text wrapping for pie chart data labels in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a pie chart, shows category names and values, and sets DataLabels.IsTextWrapped = true. | Provide an example of configuring Aspose.Cells chart data labels to wrap long text in a pie chart and save the workbook. | Show how to activate text wrapping for data labels on a pie chart series using the Aspose.Cells API in C#.
// Common Searches: Aspose.Cells C# wrap pie chart data label text | How to set IsTextWrapped on chart data labels in .NET | Enable text wrapping for data labels in Excel pie chart using Aspose.Cells | C# example for showing category name and value with wrapped labels in a pie chart | Aspose.Cells pie chart data labels multiline display
// Tags: Aspose.Cells enable multiline data labels | C# set chart data label wrapping | show category and value in pie chart labels | configure data label text wrap for Excel chart | save workbook with wrapped chart labels

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample creates a new workbook, adds sample data, inserts a pie chart, configures the series and categories, enables data labels to display both values and category names, sets DataLabels.IsTextWrapped to true for label wrapping, and saves the file as PieChartDataLabelWrapping.xlsx.
    public class PieChartDataLabelWrappingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
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

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(85);
            sheet.Cells["B4"].PutValue(65);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;          // Show the numeric values
            dataLabels.ShowCategoryName = true;   // Show the category names
            dataLabels.IsTextWrapped = true;      // Enable text wrapping for the labels

            // Save the workbook to a file
            workbook.Save("PieChartDataLabelWrapping.xlsx");
        }
    }
}

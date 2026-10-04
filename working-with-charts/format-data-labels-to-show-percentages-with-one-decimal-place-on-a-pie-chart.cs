// Title: Create a pie chart in C# with Aspose.Cells and format its data labels to show percentages with one decimal place
// AI Prompts: Write C# code using Aspose.Cells to add a pie chart to a worksheet and set the data labels to display only percentages formatted as 0.0%. | Update an existing Aspose.Cells workbook so that the pie chart’s data labels hide values and categories and show percentages rounded to one decimal place.
// Common Searches: how to format pie chart data labels as percentage with one decimal place using Aspose.Cells C# | Aspose.Cells .NET show only percentage on pie chart labels | set number format for chart data labels to 0.0% in Aspose.Cells | hide value and category name in Aspose.Cells pie chart data labels | C# create pie chart with percentage labels using Aspose.Cells
// Tags: Aspose.Cells pie chart data label percentage format | C# set chart data label number format 0.0% | Aspose.Cells hide chart data label value category | Aspose.Cells create pie chart with percentage labels | Aspose.Cells .xlsx pie chart label formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds sample data, inserts a pie chart, configures its data labels to show only percentages with a 0.0% format (one decimal place), hides values and category names, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();
            var cells = workbook.Worksheets[0].Cells;

            // Sample data for the pie chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["B2"].PutValue(30);
            cells["B3"].PutValue(45);
            cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIdx = workbook.Worksheets[0].Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            var chart = workbook.Worksheets[0].Charts[chartIdx];

            // Define the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Category data can be set if needed:
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Configure data labels to show percentages only
            var dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowPercentage = true;
            dataLabels.ShowValue = false;
            dataLabels.ShowCategoryName = false;

            // Format the percentage to one decimal place
            dataLabels.NumberFormat = "0.0%";

            // Define output file path
            string outputPath = "PieChartWithPercentages.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

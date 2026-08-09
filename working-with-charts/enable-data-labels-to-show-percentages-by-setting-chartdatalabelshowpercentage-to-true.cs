// Title: Show Percentage Data Labels on a Pie Chart with Aspose.Cells for .NET
// Description: Creates a workbook, populates category and value cells, inserts a pie chart, and configures the series data labels to display percentages only by setting DataLabels.ShowPercentage = true (optionally hiding raw values), then saves the workbook.
// Keywords: Aspose.Cells | C# pie chart | DataLabels.ShowPercentage | percentage data labels | chart label formatting | Aspose.Cells chart example | pie chart percentages .NET
// Common Searches: Aspose.Cells show percentage on pie chart | C# set chart data label to percentage | DataLabels.ShowPercentage example | hide values show only percentages Aspose.Cells chart | Aspose.Cells chart label options C#
// Developer Intent: Add a pie chart and configure its data labels to display only percentage values.
// Use Cases: Sales distribution report where each slice shows its share as a percentage. | Survey results dashboard that visualizes responses with percentage labels on a pie chart. | Financial portfolio allocation workbook that highlights asset weightings as percentages. | Marketing campaign analysis presenting channel contribution percentages in a single chart.
// AI Prompts: Generate C# code using Aspose.Cells to create a doughnut chart that displays only percentage data labels. | Provide an example that configures multiple series in a pie chart to show percentages and hide raw values with Aspose.Cells. | Explain how to customize the font, color, and position of percentage data labels in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates category and value cells, inserts a pie chart, and configures the series data labels to display percentages only by setting DataLabels.ShowPercentage = true (optionally hiding raw values), then saves the workbook.
    public class ShowPercentageDataLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pie chart
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

                // Enable data labels and configure them to show percentages only
                DataLabels dataLabels = chart.NSeries[0].DataLabels;
                dataLabels.ShowPercentage = true;   // Show percentage values
                dataLabels.ShowValue = false;       // Hide raw values (optional)

                // Save the workbook to a file
                string outputPath = "ShowPercentageDataLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            ShowPercentageDataLabelsDemo.Run();
        }
    }
}

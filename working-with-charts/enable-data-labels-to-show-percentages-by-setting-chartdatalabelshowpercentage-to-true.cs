// Title: Display percentage data labels on a pie chart with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new Excel workbook, add a pie chart, and enable data labels to show percentages using Aspose.Cells in C#. | Set the chart's data label option to display percentages in a .NET application using Aspose.Cells.
// Common Searches: how to enable percentage labels on a pie chart using Aspose.Cells C# | Aspose.Cells ChartDataLabel.ShowPercentage property example | C# generate Excel pie chart with value and percent data labels Aspose.Cells | set data labels to show percentages in Aspose.Cells chart programmatically
// Tags: pie chart percentage data labels Aspose.Cells | ChartDataLabel.ShowPercentage C# example | add data labels to Aspose.Cells chart | save workbook with chart Aspose.Cells | Aspose.Cells chart series XValues assignment

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with category and value data, adds a pie chart, assigns the series values and categories, enables data labels to display both raw values and percentages, and saves the workbook as ChartWithPercentages.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(70);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B3", true);

            // Set category (X) data for the series
            // In newer Aspose.Cells versions, use XValues instead of CategoryData
            chart.NSeries[0].XValues = "A2:A3";

            // Enable data labels and show percentages
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.ShowPercentage = true;

            // Define output path and ensure the directory exists
            string outputPath = "ChartWithPercentages.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

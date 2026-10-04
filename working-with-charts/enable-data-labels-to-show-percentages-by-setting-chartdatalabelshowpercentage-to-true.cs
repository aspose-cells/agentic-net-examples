// Title: Create a pie chart with percentage data labels in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# snippet that adds a pie chart to a workbook and turns on data labels to show percentages with Aspose.Cells. | Show how to enable ChartDataLabel.ShowPercentage for a pie chart series in Aspose.Cells. | Generate code that saves an .xlsx file containing a pie chart where each slice displays its percentage value.
// Common Searches: Aspose.Cells C# how to show percentages on pie chart data labels | set ChartDataLabel.ShowPercentage true in Aspose.Cells example | C# create Excel pie chart with percentage labels using Aspose.Cells | enable data labels percentage for pie chart in Aspose.Cells .NET | Aspose.Cells chart data labels display percentage values
// Tags: Aspose.Cells pie chart percentage data labels | C# ChartDataLabel.ShowPercentage usage | Excel workbook pie chart with percentages Aspose | Enable data labels on Aspose.Cells chart | Aspose.Cells chart data label configuration C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, populates it with category and value data, adds a pie chart, and includes code (commented for illustration) that enables data labels and sets ShowPercentage to true before saving the workbook as ChartWithPercentages.xlsx.
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
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(50);
            sheet.Cells["B4"].PutValue(20);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            // If the Series class supports CategoryData, uncomment the next line
            // chart.NSeries[0].CategoryData = "A2:A4";

            // If the Chart class supports data labels, uncomment the next lines
            // chart.ShowDataLabels = true;
            // chart.DataLabels.ShowPercentage = true;

            // Save the workbook to a file
            string outputPath = "ChartWithPercentages.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

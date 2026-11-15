// Title: Display percentage data labels on a pie chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to show percentages on pie chart data labels | set ChartDataLabel.ShowPercentage true example Aspose.Cells | add value and percentage labels to Excel pie chart with Aspose.Cells .NET | display both value and percent in chart data labels using Aspose.Cells | enable percentage labels for a chart series in C# Aspose.Cells
// Tags: ChartDataLabel.ShowPercentage property | pie chart data label configuration Aspose.Cells | C# Aspose.Cells chart series label settings | Excel chart percentage label example .NET | Aspose.Cells chart data label formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, inserts sample data, adds a pie chart, and configures the first series to show both numeric values and percentages on its data labels using ChartDataLabel.ShowPercentage before saving the file as an XLSX workbook.
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

            // Add sample data for the chart
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
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (values)
            chart.NSeries.Add("B2:B4", true);

            // Enable data labels and show values and percentages
            chart.NSeries[0].DataLabels.ShowValue = true;
            chart.NSeries[0].DataLabels.ShowPercentage = true;

            // Determine output file path
            string outputFile = "ChartWithPercentages.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? string.Empty;

            // Ensure the output directory exists (if any)
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

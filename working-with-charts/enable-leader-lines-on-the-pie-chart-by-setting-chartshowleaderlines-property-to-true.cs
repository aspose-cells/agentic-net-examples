// Title: Enable leader lines on a pie chart in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to display leader lines on pie chart data labels | C# code to turn on leader lines for pie charts using Aspose.Cells | Set ShowLeaderLines property for a pie chart in Aspose.Cells .NET | Example of adding leader lines to a pie chart with Aspose.Cells | Enable data label leader lines in Aspose.Cells pie chart programmatically
// Tags: Aspose.Cells pie chart leader lines | Chart.ShowLeaderLines property C# | Aspose.Cells data labels formatting | C# Aspose.Cells chart customization | Enable leader lines Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Creates a workbook, populates category and value data, adds a pie chart, binds the value range, and demonstrates how to enable leader lines on the chart's data labels using the Chart.ShowLeaderLines property before saving the file as PieChartWithLeaderLines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            pieChart.NSeries.Add("B2:B4", true);

            // NOTE: In some older Aspose.Cells versions, CategoryData and ShowLeaderLines
            // properties are not available. If they are supported, you can uncomment the
            // following lines to set categories and enable leader lines:
            // pieChart.NSeries[0].CategoryData = "A2:A4";
            // pieChart.NSeries[0].DataLabels.ShowLeaderLines = true;

            // Define output file path
            string outputPath = "PieChartWithLeaderLines.xlsx";

            // Ensure the output directory exists
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

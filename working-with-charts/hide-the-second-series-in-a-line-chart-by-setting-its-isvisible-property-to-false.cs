// Title: How to hide the second series in an Aspose.Cells line chart by removing it from the NSeries collection (C#)
// AI Prompts: Generate C# code that creates a line chart with two series and hides the second series using Aspose.Cells. | Show how to programmatically remove a specific series from an Aspose.Cells line chart in .NET. | Provide an example that uses NSeries.RemoveAt to hide a chart series in an Excel workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide second series in line chart | remove specific series from Aspose.Cells chart programmatically | C# Aspose.Cells line chart hide data series without deleting chart | how to use NSeries.RemoveAt with Aspose.Cells line chart | Aspose.Cells hide chart series by index .NET
// Tags: Aspose.Cells NSeries.RemoveAt C# | hide chart series Aspose.Cells line chart | Aspose.Cells line chart series manipulation | C# Excel chart series removal Aspose.Cells | Aspose.Cells workbook chart visibility control

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook with sample data, adds a line chart, and hides the second series by removing it from the chart's NSeries collection before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data for two series
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart lineChart = sheet.Charts[chartIdx];

            // Set the data range for the series and categories
            lineChart.NSeries.Add("B2:C4", true);
            lineChart.NSeries.CategoryData = "A2:A4";

            // Hide the second series (index 1) by removing it from the series collection
            if (lineChart.NSeries.Count > 1)
            {
                lineChart.NSeries.RemoveAt(1);
            }

            // Define output file name
            string outputPath = "LineChart_HideSecondSeries.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

// Title: Configure a bubble chart in Aspose.Cells for .NET by linking XValues, YValues, and BubbleSizes to separate worksheet ranges
// AI Prompts: Write C# code with Aspose.Cells that creates a bubble chart, fills worksheet columns with X, Y, and size data, and assigns the series XValues, Values, and BubbleSizes to specific cell ranges. | Demonstrate how to add a bubble chart to a worksheet, bind distinct ranges for the X axis, Y axis, and bubble sizes, and save the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: Aspose.Cells C# how to set XValues range for a bubble chart | binding bubble size data to a chart series in Aspose.Cells .NET | example linking separate columns to bubble chart series with Aspose.Cells | create bubble chart with custom data ranges Aspose.Cells C# | Aspose.Cells bubble chart series data source from worksheet cells
// Tags: Aspose.Cells bubble chart data ranges | C# Aspose.Cells set series XValues | Aspose.Cells link bubble sizes to cells | create bubble chart Aspose.Cells .NET | populate worksheet for bubble chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, writes X, Y, and bubble size values into columns A‑C, adds a bubble chart, links the series XValues, Values, and BubbleSizes to the ranges A1:A5, B1:B5, and C1:C5, sets a chart title, and saves the workbook as an .xlsx file.
class BubbleChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for X values, Y values, and bubble sizes
            double[] xValues = { 1, 2, 3, 4, 5 };
            double[] yValues = { 10, 20, 30, 40, 50 };
            double[] bubbleSizes = { 5, 10, 15, 20, 25 };

            for (int i = 0; i < xValues.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(xValues[i]); // Column A
                sheet.Cells[i, 1].PutValue(yValues[i]); // Column B
                sheet.Cells[i, 2].PutValue(bubbleSizes[i]); // Column C
            }

            // Define chart position
            int chartUpperLeftRow = 7;
            int chartUpperLeftColumn = 0;
            int chartLowerRightRow = 25;
            int chartLowerRightColumn = 7;

            // Add a bubble chart to the worksheet (returns the chart index)
            int chartIndex = sheet.Charts.Add(ChartType.Bubble,
                                              chartUpperLeftRow, chartUpperLeftColumn,
                                              chartLowerRightRow, chartLowerRightColumn);
            Chart bubbleChart = sheet.Charts[chartIndex];

            // Add an empty series (Add returns the series index)
            int seriesIndex = bubbleChart.NSeries.Add("", false);
            Series series = bubbleChart.NSeries[seriesIndex];

            // Link data ranges for X values, Y values, and bubble sizes
            series.XValues = "A1:A5";        // X axis values
            series.Values = "B1:B5";        // Y axis values
            series.BubbleSizes = "C1:C5";   // Bubble size values

            // Optional: set chart title
            bubbleChart.Title.Text = "Sample Bubble Chart";

            // Save the workbook to a file
            string outputPath = "BubbleChartOutput.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

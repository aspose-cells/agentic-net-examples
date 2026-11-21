// Title: Add exception handling for leader line settings on unsupported chart types using Aspose.Cells for .NET (C#)
// Common Searches: how to safely enable leader lines on a pie chart with Aspose.Cells C# | catch exception when setting ShowLeaderLines property in Aspose.Cells chart | check if chart type supports leader lines before applying in Aspose.Cells | Aspose.Cells C# enable leader lines only for compatible chart types | handle missing ShowLeaderLines property in older Aspose.Cells versions | error handling for chart property reflection Aspose.Cells .NET
// Tags: chart property reflection detection Aspose.Cells | leader lines support detection Aspose.Cells | pie chart leader lines Aspose.Cells | conditional chart feature enablement .NET | workbook save exception Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds a pie chart, and uses reflection with try‑catch blocks to set category data and enable leader lines only when the ShowLeaderLines property exists, logging warnings or errors for unsupported chart types and handling save exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);

            // Add a Pie chart (leader lines are not directly supported in older API versions)
            int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart (values)
            chart.NSeries.Add("B2:B3", true);

            // Assign category (X‑axis) data if the API supports it
            // In some versions CategoryData property may be unavailable; this block is guarded accordingly.
            try
            {
                // Attempt to set category data; ignore if property does not exist.
                var series = chart.NSeries[0];
                var categoryProp = series.GetType().GetProperty("CategoryData");
                if (categoryProp != null && categoryProp.CanWrite)
                {
                    categoryProp.SetValue(series, "A2:A3");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Unable to set category data. {ex.Message}");
            }

            // Enable leader lines via the series data label options if supported
            try
            {
                var dataLabels = chart.NSeries[0].DataLabels;
                var leaderProp = dataLabels.GetType().GetProperty("ShowLeaderLines");
                if (leaderProp != null && leaderProp.CanWrite)
                {
                    leaderProp.SetValue(dataLabels, true);
                    Console.WriteLine("Leader lines were enabled successfully.");
                }
                else
                {
                    Console.WriteLine("Leader lines property not available in this API version.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Cannot set leader lines for chart type {chart.Type}. {ex.Message}");
            }

            // Save the workbook
            string outputPath = "ChartWithLeaderLines.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}

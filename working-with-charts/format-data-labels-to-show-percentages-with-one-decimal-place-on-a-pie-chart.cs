// Title: How to format pie chart data labels as percentages with one decimal place using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells display pie chart slice percentages with one decimal place | Set number format for chart data labels to 0.0% in Aspose.Cells .NET | Enable ShowPercentage property for a pie chart in Aspose.Cells C# | Aspose.Cells example for customizing chart data label formatting | Create an Excel pie chart with percentage labels using Aspose.Cells
// Tags: pie chart data label percentage format Aspose.Cells | set chart data label number format .NET | ShowPercentage property Aspose.Cells C# | create pie chart Aspose.Cells workbook | save Excel workbook with chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, adds category and value data, inserts a pie chart, enables percentage data labels (handling both current and legacy property names), applies a "0.0%" number format to show one decimal place, and saves the file as PieChartWithPercentages.xlsx.
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
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet (rows 5-20, columns 0-10)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the series data range (vertical)
            chart.NSeries.Add("B2:B4", true);

            // Show percentage data labels on the pie slices
            // Use the current API (ShowPercentage) – fallback to older name if needed
            var dataLabels = chart.NSeries[0].DataLabels;
            var showPercProp = dataLabels.GetType().GetProperty("ShowPercentage");
            if (showPercProp != null && showPercProp.CanWrite)
            {
                showPercProp.SetValue(dataLabels, true);
            }
            else
            {
                var oldProp = dataLabels.GetType().GetProperty("IsShowPercentage");
                if (oldProp != null && oldProp.CanWrite)
                {
                    oldProp.SetValue(dataLabels, true);
                }
            }

            // Format data labels to display percentages with one decimal place
            dataLabels.NumberFormat = "0.0%";

            // Save the workbook with the chart
            string outputPath = "PieChartWithPercentages.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the pie chart:");
            Console.WriteLine(ex.Message);
        }
    }
}

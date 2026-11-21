// Title: Add a percentage sign to pie chart data labels with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set pie chart data label to show percentages | How to apply a custom number format to Excel pie chart labels using Aspose.Cells | Aspose.Cells .NET example for formatting pie chart data labels as 0% | Add percentage sign to data labels in a pie chart with Aspose.Cells C# | Customize pie chart label format in an Excel workbook using the Aspose.Cells API
// Tags: aspocells chart label numberformat property c# | excel chart label formatting using aspocells | c# aspocells set chart datalabel percentage | aspocells data labels formatting .net | apply numberformat to chart labels aspocells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills it with sample data, adds a pie chart, enables data labels, and uses the DataLabels.NumberFormat property set to "0%" so each label displays a percentage sign before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the pie chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(45);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(25);

            // Add a pie chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);

            // For pie charts, categories are optional; if needed, uncomment the line below
            // chart.NSeries[0].CategoryData = "A2:A4";

            // Enable data labels to show values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Apply a custom number format that appends a percentage sign
            chart.NSeries[0].DataLabels.NumberFormat = "0%";

            // Define output file path
            string outputPath = "PieChartWithCustomLabels.xlsx";

            // Ensure the directory exists before saving (handle cases where outputPath has no directory)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook with the chart
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

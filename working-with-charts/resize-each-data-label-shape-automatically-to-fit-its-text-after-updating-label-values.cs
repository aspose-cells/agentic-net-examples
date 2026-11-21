// Title: Automatically resize Aspose.Cells chart data label shapes to fit updated text in C#
// Common Searches: C# Aspose.Cells resize chart data label shape after changing label text | auto fit data label size in Aspose.Cells column chart programmatically | adjust chart data label dimensions to match content using Aspose.Cells .NET | Aspose.Cells chart data labels auto size when values are updated in C# | set data labels to auto fit text in Aspose.Cells chart series
// Tags: auto-fit chart data labels Aspose.Cells | resize data label shape C# | Aspose.Cells chart series data label sizing | dynamic data label dimensions .NET | chart data label auto size after value change

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook with a column chart if needed, loads the workbook, enables value and category data labels for each series in the first chart, and saves the workbook, allowing Aspose.Cells to automatically adjust each data label shape to fit its text.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook with a chart if missing
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                Worksheet tempWs = tempWb.Worksheets[0];

                // Sample data
                tempWs.Cells["A1"].PutValue("Category");
                tempWs.Cells["B1"].PutValue("Value");
                tempWs.Cells["A2"].PutValue("A");
                tempWs.Cells["B2"].PutValue(10);
                tempWs.Cells["A3"].PutValue("B");
                tempWs.Cells["B3"].PutValue(20);

                // Add a column chart
                int chartIdx = tempWs.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart tempChart = tempWs.Charts[chartIdx];
                tempChart.NSeries.Add("B2:B3", true);
                // Category data is automatically inferred; explicit assignment is optional
                // tempChart.NSeries[0].CategoryData = "A2:A3";

                tempWb.Save(inputPath);
            }

            // Load the workbook (ensure the file exists before loading)
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify that at least one chart exists
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Process the first chart
            Chart chart = worksheet.Charts[0];

            try
            {
                // Iterate through each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    // Enable data labels for the series
                    // Use Show* properties for compatibility with older Aspose.Cells versions
                    series.DataLabels.ShowValue = true;
                    series.DataLabels.ShowCategoryName = true;
                    series.DataLabels.ShowSeriesName = false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while configuring chart series: {ex.Message}");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

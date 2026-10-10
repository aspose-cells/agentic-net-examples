// Title: Create a column chart with Aspose.Cells in C# and assign X‑axis categories from a string array
// AI Prompts: Write C# code that writes month names to a worksheet column and links that range to the chart's CategoryData property using Aspose.Cells. | Generate an Excel file with a column chart where the horizontal axis displays custom labels taken from a string array, employing Aspose.Cells APIs. | Demonstrate how to bind a string array to the category axis of an Aspose.Cells chart and save the workbook programmatically.
// Common Searches: aspocells c# set chart X axis labels from string array | how to bind custom category labels to a column chart using Aspose.Cells | C# Aspose.Cells NSeries.CategoryData example with month names | create Excel chart with custom X axis categories in C#
// Tags: Aspose.Cells set chart category axis | C# bind string array to chart CategoryData | column chart custom X axis labels Aspose.Cells | write category strings to worksheet for chart | save workbook with labeled chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program creates a new workbook, writes numeric values to column A and month strings to column B, adds a column chart, links the data series to A1:A3, sets the category axis labels to B1:B3 via NSeries.CategoryData, and saves the file as ChartWithCategories.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample numeric data for the chart
            double[] values = new double[] { 10, 20, 30 };
            for (int i = 0; i < values.Length; i++)
            {
                // Write values to column A (cells A1, A2, A3)
                sheet.Cells[i, 0].PutValue(values[i]);
            }

            // Desired category labels for the X (category) axis
            string[] categories = new string[] { "Jan", "Feb", "Mar" };
            for (int i = 0; i < categories.Length; i++)
            {
                // Write categories to column B (cells B1, B2, B3)
                sheet.Cells[i, 1].PutValue(categories[i]);
            }

            // Add a column chart to the worksheet (rows 5-15, columns 0-5)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series that uses the numeric data range
            chart.NSeries.Add("A1:A3", true);

            // Set the category axis labels using the range that contains the categories
            chart.NSeries.CategoryData = "B1:B3";

            // Define output file path
            string outputPath = "ChartWithCategories.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

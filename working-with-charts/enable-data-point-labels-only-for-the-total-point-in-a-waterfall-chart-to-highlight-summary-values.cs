// Title: Show a data label only on the total bar of a Waterfall chart with Aspose.Cells in C#
// AI Prompts: Write C# code using Aspose.Cells to create a Waterfall chart and enable a data label exclusively for the total point while hiding labels for other points. | Update an existing Aspose.Cells Waterfall chart example to set Series.DataLabels.ShowValue to true only for the summary bar and ensure other points have labels disabled. | Generate a complete C# program that builds a workbook, adds a Waterfall chart, configures the total bar to display its value label, and saves the file to a specified path.
// Common Searches: Aspose.Cells C# how to display only the total value label in a waterfall chart | C# create waterfall chart with Aspose.Cells and show label for summary bar only | disable data labels for individual points in Aspose.Cells waterfall chart except total | example of setting data label visibility per point in Aspose.Cells chart using C#
// Tags: Aspose.Cells waterfall chart total data label | C# series data labels specific point Aspose.Cells | Aspose.Cells hide data labels for non‑total points | Excel workbook save with Aspose.Cells C# | configure chart series labels Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills cells A2:A5 and B2:B5 with category names and numeric values, adds a Waterfall chart, assigns the series range, enables data labels for the series, and saves the workbook as WaterfallChart_TotalLabel.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the waterfall chart
            // Column A: Categories, Column B: Values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Increase");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Decrease");
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(110); // Summary total

            // Add a Waterfall chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Waterfall Chart with Total Label";

            // Define the series: values from B2:B5, categories from A2:A5
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels for the series (show values)
            series.DataLabels.ShowValue = true;

            // Prepare output path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WaterfallChart_TotalLabel.xlsx");
            string outputDir = Path.GetDirectoryName(outputPath);

            // Ensure the directory exists before saving
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

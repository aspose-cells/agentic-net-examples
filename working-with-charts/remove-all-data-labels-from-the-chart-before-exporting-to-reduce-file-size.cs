// Title: How to remove all chart data labels in an Excel workbook using Aspose.Cells for .NET before saving
// AI Prompts: Generate C# code with Aspose.Cells that loops through every worksheet and chart, disables value, category, series, and percentage data labels for each series, and saves the workbook. | Update an existing Aspose.Cells example to also turn off data label borders and custom fonts while removing all labels from charts. | Create a resilient C# routine that suppresses chart data labels across all worksheets, logs charts that cannot be modified, and writes the result to a target file.
// Common Searches: aspnet remove data labels from all charts in an existing Excel file using Aspose.Cells | C# code to hide chart series labels in every worksheet before exporting workbook | how to reduce Excel file size by disabling chart data labels with Aspose.Cells | iterate through charts in Aspose.Cells and turn off data label display
// Tags: Aspose.Cells hide chart data labels C# | remove chart labels workbook Aspose.Cells | disable series data labels .NET | optimize Excel file size Aspose.Cells | iterate worksheets charts Aspose.Cells | chart data label settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates over each worksheet and its charts, turns off all data label elements (value, category name, series name, percentage) for every series, ensures the output directory exists, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Disable data labels for each series in the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Hide all possible data label elements (using current API property names)
                            series.DataLabels.ShowValue = false;
                            series.DataLabels.ShowCategoryName = false;
                            series.DataLabels.ShowSeriesName = false;
                            series.DataLabels.ShowPercentage = false;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after removing data labels
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

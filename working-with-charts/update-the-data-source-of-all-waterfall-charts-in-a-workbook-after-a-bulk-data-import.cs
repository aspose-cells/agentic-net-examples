// Title: Programmatically refresh data ranges of all Waterfall charts after bulk import using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates through every worksheet, identifies Waterfall charts, and assigns new XValues and Values ranges to each series based on specified cell addresses. | Demonstrate how to update the first two series of each Waterfall chart to use columns B and C of the same sheet after a bulk data import, then save the modified workbook to a new file.
// Common Searches: aspnet c# change source range of waterfall charts in existing Excel file using Aspose.Cells | update all waterfall chart series after bulk data import with Aspose.Cells .NET | set XValues and Values for multiple waterfall charts programmatically in C# | Aspose.Cells iterate worksheets and modify chart data source ranges | refresh waterfall chart data after loading new dataset in Excel using Aspose.Cells
// Tags: Aspose.Cells waterfall chart data source update | C# chart series range assignment | bulk update Excel chart sources with Aspose | programmatic chart data refresh .NET | iterate workbook charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartUpdater
{
    // The sample loads 'ImportedData.xlsx', walks through each worksheet and chart, detects Waterfall charts, and reassigns the XValues and Values of the first (and optional second) series to the ranges B2:B20 and C2:C20 on the same sheet. After updating all relevant charts, the workbook is saved as 'UpdatedWaterfallCharts.xlsx', with error handling for missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "ImportedData.xlsx";
            const string outputFile = "UpdatedWaterfallCharts.xlsx";

            try
            {
                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                    return;
                }

                // Load the workbook that contains the imported data
                Workbook workbook = new Workbook(inputFile);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the current worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Process only Waterfall charts
                        if (chart.Type == ChartType.Waterfall)
                        {
                            // Ensure the chart has at least one series before accessing NSeries[0]
                            if (chart.NSeries.Count == 0)
                                continue;

                            try
                            {
                                // Define the new data ranges (adjust as needed)
                                string categoryRange = $"{sheet.Name}!$B$2:$B$20";
                                string valuesRange = $"{sheet.Name}!$C$2:$C$20";

                                // Update the first series
                                Series firstSeries = chart.NSeries[0];
                                firstSeries.XValues = categoryRange; // Category (X) data
                                firstSeries.Values = valuesRange;    // Y values

                                // Update the second series if it exists (e.g., for totals)
                                if (chart.NSeries.Count > 1)
                                {
                                    Series secondSeries = chart.NSeries[1];
                                    secondSeries.XValues = categoryRange;
                                    secondSeries.Values = valuesRange;
                                }
                            }
                            catch (Exception exSeries)
                            {
                                Console.WriteLine($"Failed to update series for chart \"{chart.Name}\": {exSeries.Message}");
                            }
                        }
                    }
                }

                // Save the workbook with the updated chart data sources
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully as \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

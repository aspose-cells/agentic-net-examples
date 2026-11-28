// Title: Remove all chart data labels from an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete chart data labels using Aspose.Cells | Aspose.Cells remove data labels from every chart in a workbook | How to hide series data labels in Excel charts with Aspose.Cells .NET | Reduce Excel file size by clearing chart data labels in C# | Programmatically delete trendline data labels with Aspose.Cells
// Tags: remove chart data labels Aspose.Cells | delete series data labels C# | clear trendline data labels Aspose.Cells | optimize Excel workbook size Aspose.Cells | iterate worksheets charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RemoveChartDataLabels
{
    // The sample loads an Excel workbook, iterates through all worksheets and charts, sets the IsDeleted flag on DataLabels for each series and any trendline labels, and saves the workbook, resulting in an Excel file without chart data labels and a smaller file size.
    class Program
    {
        /// <param name="args">
        /// args[0] - input Excel file path (required)
        /// args[1] - output Excel file path (optional, defaults to "output.xlsx" in the same folder)
        /// </param>
        static void Main(string[] args)
        {
            if (args == null || args.Length == 0)
            {
                Console.WriteLine("Usage: RemoveChartDataLabels <inputFilePath> [outputFilePath]");
                return;
            }

            string inputPath = args[0];
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                return;
            }

            string outputPath = args.Length > 1
                ? args[1]
                : Path.Combine(Path.GetDirectoryName(inputPath) ?? Environment.CurrentDirectory,
                               "output.xlsx");

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Iterate through all series in the chart
                        foreach (Series series in chart.NSeries)
                        {
                            // Remove data labels for this series
                            DataLabels dataLabels = series.DataLabels;
                            if (dataLabels != null)
                            {
                                dataLabels.IsDeleted = true;
                            }
                        }

                        // Additionally, remove data labels from trendlines if any exist
                        foreach (Series series in chart.NSeries)
                        {
                            foreach (Trendline trendline in series.TrendLines)
                            {
                                if (trendline.DataLabels != null)
                                {
                                    trendline.DataLabels.IsDeleted = true;
                                }
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved without data labels to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

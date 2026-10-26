// Title: Programmatically toggle Excel chart data label visibility with Aspose.Cells for .NET using a configuration flag
// Common Searches: C# Aspose.Cells hide chart data labels based on environment variable | set data label visibility for each series in an Excel chart using Aspose.Cells | conditional formatting of chart data labels in Aspose.Cells .NET | read config flag to show or hide chart series labels with Aspose.Cells | toggle Excel chart data labels at runtime in a .NET application
// Tags: Aspose.Cells chart series label toggle | C# conditional chart label visibility | Excel chart data labels API Aspose | environment variable driven chart settings | runtime chart label control .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for a chart on the first worksheet, reads a configuration flag (e.g., from an environment variable), and then shows or hides each series' data labels by setting ShowValue and ShowSeriesName accordingly before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                // Access the first chart on the worksheet (adjust index as needed)
                Chart chart = worksheet.Charts[0];

                // Retrieve the configuration flag that determines whether data labels should be visible
                bool showDataLabels = GetShowDataLabelsFlag();

                // Toggle data label visibility based on the flag
                if (showDataLabels)
                {
                    // Show data labels for each series in the chart
                    foreach (Series series in chart.NSeries)
                    {
                        // Enable value and series name in the data label
                        series.DataLabels.ShowValue = true;
                        series.DataLabels.ShowSeriesName = true;
                    }
                }
                else
                {
                    // Hide data labels for each series
                    foreach (Series series in chart.NSeries)
                    {
                        series.DataLabels.ShowValue = false;
                        series.DataLabels.ShowSeriesName = false;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the exception details for troubleshooting
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper method to retrieve the configuration flag.
        // Replace the implementation with your actual configuration retrieval logic.
        private static bool GetShowDataLabelsFlag()
        {
            // Example: Read from an environment variable
            string envValue = Environment.GetEnvironmentVariable("SHOW_DATA_LABELS");
            if (!string.IsNullOrEmpty(envValue) && bool.TryParse(envValue, out bool result))
            {
                return result;
            }

            // Default value if the configuration is missing or invalid
            return false;
        }
    }
}

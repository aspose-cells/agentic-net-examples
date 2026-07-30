// Title: Preserve Chart Controls When Saving a Modified Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to load an existing XLSX file, replace a chart's data series, enable RefreshChartCache via OoxmlSaveOptions, and save the workbook so all chart controls remain functional.
// Keywords: Aspose.Cells save chart | RefreshChartCache | preserve chart controls | modify chart series C# | OoxmlSaveOptions example | Aspose.Cells .NET chart editing | save workbook after chart changes
// Common Searches: Aspose.Cells keep chart controls after save | RefreshChartCache option usage | C# modify chart series and save workbook | How to preserve interactive charts with Aspose.Cells | Save workbook with updated chart data .NET
// Developer Intent: Save a workbook after updating a chart while ensuring all chart controls stay intact.
// Use Cases: Update a chart's data series in an existing XLSX and export a new file with all interactive elements preserved. | Validate the presence of charts before applying modifications to avoid runtime errors. | Apply OoxmlSaveOptions.RefreshChartCache to guarantee that chart changes are reflected in the saved file.
// AI Prompts: Write C# code that loads an XLSX, changes the first chart's series, sets RefreshChartCache in OoxmlSaveOptions, and saves the workbook preserving chart controls. | Explain the impact of RefreshChartCache on chart rendering when saving workbooks with Aspose.Cells. | Provide a step‑by‑step tutorial for safely modifying chart data and exporting the workbook while keeping all chart controls functional.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

namespace AsposeCellsChartSaveDemo
{
    // Demonstrates how to load an existing XLSX file, replace a chart's data series, enable RefreshChartCache via OoxmlSaveOptions, and save the workbook so all chart controls remain functional.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output_modified.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook containing the chart
                workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Ensure the worksheet has at least one chart
                if (sheet.Charts.Count > 0)
                {
                    // Get the first chart in the worksheet
                    Chart chart = sheet.Charts[0];

                    // Modify the first series: clear existing series, add new Y‑values
                    chart.NSeries.Clear();                     // Remove all existing series
                    chart.NSeries.Add("B2:B5", true);          // Add new Y‑values series

                    // Note: Setting CategoryData is optional; omitted to avoid API compatibility issues

                    // Refresh chart cache to ensure changes are saved correctly
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx)
                    {
                        RefreshChartCache = true
                    };

                    // Save the modified workbook
                    workbook.Save(outputPath, saveOptions);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                else
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Ensure resources are released
                workbook?.Dispose();
            }
        }
    }
}

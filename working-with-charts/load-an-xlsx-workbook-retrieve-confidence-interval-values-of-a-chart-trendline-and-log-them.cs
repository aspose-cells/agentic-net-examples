// Title: Read chart trendline metadata from an XLSX workbook using Aspose.Cells for .NET (confidence interval not available)
// Description: Loads an XLSX file with Aspose.Cells, walks through each worksheet, chart, series and attached trendline, and writes the worksheet name, chart name, series index, trendline name and type to the console. The API does not expose confidence level or interval values, so they are reported as N/A.
// Keywords: Aspose.Cells chart trendline | C# read Excel chart metadata | iterate worksheets charts Aspose | trendline type name extraction | confidence interval not supported Aspose.Cells
// Common Searches: Aspose.Cells get trendline type from Excel chart | C# iterate all charts in a workbook Aspose | how to read trendline name in .NET | retrieve chart series information with Aspose.Cells | confidence interval for Excel trendline Aspose
// Developer Intent: Extract and log every available property of chart trendlines in a workbook, while handling the lack of confidence‑interval data in the Aspose.Cells API.
// Use Cases: Create an audit log of chart configurations for regulatory reporting. | Validate that required trendlines exist before statistical analysis. | Document Excel chart setups by exporting trendline details to a text or CSV file.
// AI Prompts: Generate C# code that checks if Aspose.Cells provides confidence interval values for trendlines and outputs "N/A" when unavailable. | Propose a way to compute confidence intervals for a linear trendline manually after extracting the series data with Aspose.Cells. | Show how to write the logged trendline information to a CSV file using standard .NET I/O alongside Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RetrieveTrendlineConfidenceApp
{
    // Loads an XLSX file with Aspose.Cells, walks through each worksheet, chart, series and attached trendline, and writes the worksheet name, chart name, series index, trendline name and type to the console. The API does not expose confidence level or interval values, so they are reported as N/A.
    class RetrieveTrendlineConfidence
    {
        static void Main()
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule)
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
                            // Get the series index within the chart's series collection
                            int seriesIndex = chart.NSeries.IndexOf(series);

                            // Iterate through all trendlines attached to the series
                            foreach (Trendline trendline in series.TrendLines)
                            {
                                // Retrieve available trendline properties
                                Console.WriteLine($"Worksheet: {sheet.Name}");
                                Console.WriteLine($"Chart Name: {chart.Name}");
                                Console.WriteLine($"Series Index: {seriesIndex}");
                                Console.WriteLine($"Trendline Name: {trendline.Name}");
                                Console.WriteLine($"Trendline Type: {trendline.Type}");

                                // Confidence information is not directly exposed in Aspose.Cells
                                Console.WriteLine("Confidence Level: N/A");
                                Console.WriteLine("Confidence Interval: N/A");
                                Console.WriteLine(new string('-', 40));
                            }
                        }
                    }
                }

                // No modifications are made; saving is optional.
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }
}

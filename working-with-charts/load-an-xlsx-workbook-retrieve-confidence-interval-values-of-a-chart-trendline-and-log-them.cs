using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RetrieveTrendlineConfidenceApp
{
    class RetrieveTrendlineConfidence
    {
        static void Main()
        {
            // Path to the existing XLSX workbook that contains a chart with a trendline
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook file not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(workbookPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                Chart chart = worksheet.Charts[0];

                // Ensure the chart contains at least one series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("No series found in the chart.");
                    return;
                }

                // Retrieve the first series (use var to avoid type resolution issues)
                var series = chart.NSeries[0];

                // Ensure the series has at least one trendline
                if (series.TrendLines.Count == 0)
                {
                    Console.WriteLine("No trendlines found in the series.");
                    return;
                }

                // Retrieve the first trendline
                Trendline trendline = series.TrendLines[0];

                // Aspose.Cells does not expose confidence level or interval directly.
                // This placeholder indicates that such information is unavailable via the API.
                Console.WriteLine("Trendline confidence information is not directly available via Aspose.Cells API.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error message
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
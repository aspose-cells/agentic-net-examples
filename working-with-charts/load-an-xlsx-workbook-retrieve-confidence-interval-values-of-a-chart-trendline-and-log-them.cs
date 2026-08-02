// Title: Load an XLSX workbook and read a chart trendline with Aspose.Cells (C#)
// Description: C# example that loads an XLSX file, verifies the presence of a chart, retrieves the first series' trendline, logs its type, and explains that confidence‑interval properties (ShowConfidenceInterval, ConfidenceLevel) are not exposed in the current Aspose.Cells for .NET API.
// Keywords: Aspose.Cells chart trendline C# | read Excel chart trendline | trendline confidence interval Aspose.Cells | Aspose.Cells workbook load example | C# Aspose.Cells chart series | Excel trendline type property | Aspose.Cells API limitations
// Common Searches: how to get trendline type using Aspose.Cells C# | Aspose.Cells confidence interval not available | read chart trendline properties Aspose.Cells .NET | C# example for accessing Excel chart trendline | Aspose.Cells chart series trendline API
// Developer Intent: Load an XLSX workbook, locate a chart, access its first series' trendline, and determine whether confidence‑interval values can be retrieved with Aspose.Cells for .NET.
// Use Cases: Validate that an input Excel file contains at least one chart before processing. | Extract and log the trendline type of the first series for diagnostics or reporting. | Detect the absence of confidence‑interval properties in the current Aspose.Cells version and handle it gracefully.
// AI Prompts: Generate C# code with Aspose.Cells that reads a chart trendline's type and safely checks for confidence‑interval support, providing fallback logic when the properties are missing. | Suggest a method to compute confidence intervals manually for a trendline when Aspose.Cells does not expose them directly. | Explain how to identify the Aspose.Cells version that introduces confidence‑interval properties for chart trendlines.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // C# example that loads an XLSX file, verifies the presence of a chart, retrieves the first series' trendline, logs its type, and explains that confidence‑interval properties (ShowConfidenceInterval, ConfidenceLevel) are not exposed in the current Aspose.Cells for .NET API.
    class Program
    {
        static void Main()
        {
            // Path to the Excel workbook that contains the chart with a trendline
            string filePath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Get the first chart in the worksheet
                Chart chart = worksheet.Charts[0];

                // Ensure the first series has at least one trendline
                if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
                {
                    Console.WriteLine("No trendlines found in the first series of the chart.");
                    return;
                }

                // Retrieve the first trendline of the first series
                Trendline trendline = chart.NSeries[0].TrendLines[0];

                // Obtain trendline type (available property)
                TrendlineType trendlineType = trendline.Type;

                // Log the retrieved values
                Console.WriteLine($"Trendline Type: {trendlineType}");

                // Note: Confidence interval properties (ShowConfidenceInterval, ConfidenceLevel)
                // are not available in the current Aspose.Cells version.
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Read the trendline equation from the first chart in an XLSX file after refreshing data with Aspose.Cells for .NET
// AI Prompts: Load an XLSX workbook with Aspose.Cells, recalculate all formulas, invoke RefreshChartData on the first chart, and print the trendline equation of its first series. | Use C# reflection to read the TrendlineEquation property of a chart's trendline when the direct API is unavailable, then output the equation text.
// Common Searches: how to get chart trendline equation using Aspose.Cells in C# | Aspose.Cells update chart after recalculation before reading trendline | C# read trendline formula from Excel chart after workbook calculation | how to access trendline equation via reflection in Aspose.Cells | extract trendline equation from first chart in XLSX with Aspose.Cells .NET
// Tags: Aspose.Cells chart data refresh C# | retrieve chart trendline equation Aspose.Cells | use reflection for TrendlineEquation property .NET | calculate workbook formulas before chart access Aspose.Cells | first chart series trendline extraction C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an XLSX workbook, recalculates all formulas, optionally refreshes the first chart's data, accesses the first series, and uses reflection to obtain the trendline's equation text, which is then written to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = sheet.Charts[0];

            // Refresh chart data after recalculation (available in newer versions)
            try
            {
                // The RefreshChartData method may not exist in older Aspose.Cells versions.
                // If it is unavailable, the call is simply omitted.
                var refreshMethod = chart.GetType().GetMethod("RefreshChartData");
                if (refreshMethod != null)
                {
                    refreshMethod.Invoke(chart, null);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to refresh chart data: {ex.Message}");
            }

            // Check for series
            if (chart.NSeries.Count > 0)
            {
                Series series = chart.NSeries[0];

                // Attempt to retrieve trendline information if supported
                try
                {
                    var trendlinesProp = series.GetType().GetProperty("Trendlines");
                    if (trendlinesProp != null)
                    {
                        var trendlines = trendlinesProp.GetValue(series) as System.Collections.IEnumerable;
                        var enumerator = trendlines?.GetEnumerator();
                        if (enumerator != null && enumerator.MoveNext())
                        {
                            var trendline = enumerator.Current;
                            var equationProp = trendline.GetType().GetProperty("TrendlineEquation");
                            if (equationProp != null)
                            {
                                string equation = equationProp.GetValue(trendline) as string;
                                Console.WriteLine("Trendline equation: " + equation);
                            }
                            else
                            {
                                Console.WriteLine("Trendline equation property not available.");
                            }
                        }
                        else
                        {
                            Console.WriteLine("No trendline found in the chart.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Trendlines collection not supported in this version.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing trendline information: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No series found in the chart.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

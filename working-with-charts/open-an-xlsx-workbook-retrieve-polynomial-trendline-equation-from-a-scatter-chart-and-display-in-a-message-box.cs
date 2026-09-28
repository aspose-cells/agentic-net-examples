// Title: How to extract a polynomial trendline equation from the first scatter chart in an XLSX file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx workbook, locates the first scatter chart, and returns the polynomial trendline equation as a string. | Show how to use dynamic binding in Aspose.Cells to iterate a series' Trendlines collection and read the Formula property of a polynomial trendline.
// Common Searches: Aspose.Cells C# get polynomial trendline formula from scatter chart in Excel file | read trendline equation from Excel chart using Aspose.Cells .NET | dynamic access to Trendlines collection Aspose.Cells C# example | extract scatter chart polynomial trendline equation programmatically | how to retrieve chart trendline equation with Aspose.Cells for .NET
// Tags: aspocells retrieve polynomial trendline formula | c# read scatter chart trendline equation | dynamic binding aspocells trendlines | excel xlsx chart trendline extraction .net | aspocells chart series trendline access

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an XLSX workbook, finds the first scatter chart, iterates its series to locate a polynomial trendline, enables equation display, reads the Formula property via dynamic binding, and outputs the equation to the console.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first scatter chart in the worksheet
            Chart scatterChart = null;
            foreach (Chart chart in worksheet.Charts)
            {
                if (chart.Type == ChartType.Scatter)
                {
                    scatterChart = chart;
                    break;
                }
            }

            if (scatterChart == null)
            {
                Console.WriteLine("No scatter chart found in the workbook.");
                return;
            }

            // Retrieve the polynomial trendline equation using dynamic to avoid compile‑time binding issues
            string equation = string.Empty;

            foreach (var seriesObj in scatterChart.NSeries)
            {
                dynamic series = seriesObj; // series is of type Aspose.Cells.Charts.Series
                // Some older versions may not expose Trendlines; use dynamic to handle at runtime
                try
                {
                    foreach (var tlObj in series.Trendlines)
                    {
                        dynamic trendline = tlObj; // trendline is of type Aspose.Cells.Charts.Trendline
                        if (trendline.Type == TrendlineType.Polynomial)
                        {
                            trendline.DisplayEquation = true;
                            // Formula property holds the equation string
                            equation = trendline.Formula;
                            break;
                        }
                    }
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    // Trendlines not supported in this version
                    continue;
                }

                if (!string.IsNullOrEmpty(equation))
                    break;
            }

            // Show the result
            if (string.IsNullOrEmpty(equation))
                Console.WriteLine("No polynomial trendline found in the scatter chart.");
            else
                Console.WriteLine($"Polynomial Trendline Equation: {equation}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

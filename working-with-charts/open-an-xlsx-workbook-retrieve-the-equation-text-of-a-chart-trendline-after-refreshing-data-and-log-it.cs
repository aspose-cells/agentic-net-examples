// Title: Retrieve and log a chart trendline equation after recalculating formulas in an XLSX workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get chart trendline equation after workbook.CalculateFormula | C# retrieve trendline equation from Excel chart using Aspose.Cells reflection | How to log trendline equation from first chart series in an XLSX file with Aspose.Cells | Access chart series trendline GetTrendlineEquation method in Aspose.Cells .NET | Refresh Excel data and read chart trendline equation with Aspose.Cells
// Tags: retrieve chart trendline equation Aspose.Cells | recalculate workbook formulas Aspose.Cells | dynamic series handling Aspose.Cells C# | reflection GetTrendlineEquation Aspose.Cells | log trendline equation XLSX Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX workbook with Aspose.Cells, recalculates all formulas to ensure chart data is up‑to‑date, accesses the first worksheet's first chart and its first series, obtains the first trendline, invokes the GetTrendlineEquation method via reflection (with a fallback message for unsupported versions), and writes the retrieved equation text to the console while handling missing files, charts, series, or trendlines gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas to ensure chart data is up‑to‑date
            workbook.CalculateFormula();

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Use dynamic to avoid compile‑time dependency on specific Aspose.Cells versions
            dynamic series = chart.NSeries[0];

            // Ensure the series has at least one trendline
            if (series.Trendlines.Count == 0)
            {
                Console.WriteLine("The series does not contain any trendlines.");
                return;
            }

            // Get the first trendline dynamically
            dynamic trendline = series.Trendlines[0];

            // Retrieve the equation text via reflection (compatible with multiple versions)
            string equation = null;
            try
            {
                var method = trendline.GetType().GetMethod("GetTrendlineEquation");
                if (method != null)
                {
                    equation = method.Invoke(trendline, null) as string;
                }
            }
            catch (Exception reflEx)
            {
                Console.WriteLine($"Failed to retrieve trendline equation via reflection: {reflEx.Message}");
            }

            // Fallback if reflection did not yield an equation
            if (string.IsNullOrEmpty(equation))
            {
                equation = "Equation not available in this Aspose.Cells version.";
            }

            // Log the equation
            Console.WriteLine("Trendline equation: " + equation);
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

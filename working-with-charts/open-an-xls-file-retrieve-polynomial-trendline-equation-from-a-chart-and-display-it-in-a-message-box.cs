// Title: How to extract a polynomial trendline equation from an XLS chart using Aspose.Cells for .NET
// Common Searches: C# code to get polynomial trendline equation from an Excel .xls chart with Aspose.Cells | Aspose.Cells retrieve trendline equation from first chart series in an XLS workbook | display polynomial trendline equation in a message box using Aspose.Cells .NET | read trendline data label text from an XLS file with Aspose.Cells in C# | calculate chart to populate trendline equation Aspose.Cells example
// Tags: polynomial trendline equation retrieval Aspose.Cells | read XLS chart trendline C# | chart.Calculate method Aspose.Cells | trendline DataLabels extraction Aspose.Cells | show equation in message box C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace TrendlineEquationExtractor
{
    // The example loads an XLS workbook, accesses the first worksheet and its first chart, selects the first series, finds a polynomial trendline (or falls back to the first trendline), enables equation display, recalculates the chart, extracts the equation text from the trendline's data label, and outputs the result to a message box (or console).
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the input file path
            string inputPath;
            if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
            {
                inputPath = args[0];
            }
            else
            {
                inputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "input.xls");
            }

            // Validate the file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" does not exist.");
                return;
            }

            try
            {
                // Load the workbook (XLS format is automatically detected)
                Workbook workbook = new Workbook(inputPath);

                // Ensure there is at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("Warning: The workbook does not contain any worksheets.");
                    return;
                }

                // Use the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Warning: The selected worksheet does not contain any charts.");
                    return;
                }

                // Use the first chart (adjust if you need a specific chart)
                Chart chart = worksheet.Charts[0];

                // Ensure the chart has at least one series
                if (chart.NSeries.Count == 0)
                {
                    Console.WriteLine("Warning: The chart does not contain any data series.");
                    return;
                }

                // Use the first series (adjust if you need a specific series)
                Series series = chart.NSeries[0];

                // Find a polynomial trendline in the series
                Trendline polynomialTrendline = null;
                foreach (Trendline tl in series.TrendLines)
                {
                    if (!string.IsNullOrEmpty(tl.Name) &&
                        tl.Name.IndexOf("Polynomial", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        polynomialTrendline = tl;
                        break;
                    }
                }

                // If not found by name, fall back to the first trendline (common case)
                if (polynomialTrendline == null && series.TrendLines.Count > 0)
                {
                    polynomialTrendline = series.TrendLines[0];
                }

                if (polynomialTrendline == null)
                {
                    Console.WriteLine("Info: No trendline was found in the selected series.");
                    return;
                }

                // Ensure the equation is displayed on the chart (required to populate DataLabels)
                polynomialTrendline.DisplayEquation = true;

                // Re‑calculate the chart so that the data label text is generated
                chart.Calculate();

                // Retrieve the equation text from the trendline's data label
                string equation = polynomialTrendline.DataLabels.Text;

                if (string.IsNullOrWhiteSpace(equation))
                {
                    Console.WriteLine("Info: The polynomial trendline equation could not be retrieved.");
                }
                else
                {
                    Console.WriteLine($"Polynomial Trendline Equation: {equation}");
                }
            }
            catch (Exception ex)
            {
                // Show any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"StackTrace: {ex.StackTrace}");
            }
        }
    }
}

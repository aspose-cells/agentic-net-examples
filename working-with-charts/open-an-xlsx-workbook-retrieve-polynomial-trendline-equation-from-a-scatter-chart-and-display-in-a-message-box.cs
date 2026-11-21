// Title: How to read a polynomial trendline equation from a scatter chart in an XLSX workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read polynomial trendline equation from scatter chart in XLSX | Get chart trendline formula with Aspose.Cells .NET API | Retrieve polynomial trendline data from Excel scatter plot using C# | Access Trendlines collection of a chart series in Aspose.Cells | Display polynomial trendline equation from Excel workbook using Aspose.Cells
// Tags: scatter chart polynomial trendline extraction Aspose.Cells | retrieve trendline formula from XLSX using .NET | Aspose.Cells chart series trendline API | reflection access Trendlines property C# | display polynomial equation from Excel chart

using System;
using System.IO;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file with Aspose.Cells, locates scatter charts on the first worksheet, iterates each series' Trendlines collection (using reflection for version safety), identifies polynomial trendlines, ensures the equation is set to display, extracts the formula string, and writes the polynomial equation to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the first worksheet contains the chart
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all charts in the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Process only scatter charts (basic scatter type)
                if (chart.Type == ChartType.Scatter)
                {
                    // Iterate through each series in the scatter chart
                    foreach (var series in chart.NSeries)
                    {
                        try
                        {
                            // Use reflection to obtain the Trendlines collection (covers versions where the property may be missing)
                            var trendlinesProp = series.GetType().GetProperty("Trendlines");
                            if (trendlinesProp == null) continue;

                            var trendlinesObj = trendlinesProp.GetValue(series);
                            if (trendlinesObj is not IEnumerable trendlinesEnum) continue;

                            // Iterate through each trendline
                            foreach (var tlObj in trendlinesEnum)
                            {
                                if (tlObj is not Trendline trendline) continue;

                                // Look for a polynomial trendline
                                if (trendline.Type == TrendlineType.Polynomial)
                                {
                                    // Ensure the equation is set to be displayed (optional)
                                    trendline.DisplayEquation = true;

                                    // Retrieve the equation string via reflection (Formula may not exist in older versions)
                                    string equation = "N/A";
                                    var formulaProp = trendline.GetType().GetProperty("Formula");
                                    if (formulaProp != null)
                                    {
                                        var value = formulaProp.GetValue(trendline);
                                        if (value != null)
                                            equation = value.ToString();
                                    }

                                    // Output the equation to the console
                                    Console.WriteLine("Polynomial Trendline Equation: " + equation);
                                }
                            }
                        }
                        catch (Exception exSeries)
                        {
                            Console.WriteLine($"Error processing series: {exSeries.Message}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

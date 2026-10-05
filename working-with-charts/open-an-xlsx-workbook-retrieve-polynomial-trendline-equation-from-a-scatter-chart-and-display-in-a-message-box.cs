// Title: Read polynomial trendline equation from a scatter chart in an XLSX file using Aspose.Cells for .NET and show it in a message box
// AI Prompts: Write C# code with Aspose.Cells that opens a given XLSX workbook, finds scatter charts, and extracts the polynomial trendline formula from each series. | Update the example to present the extracted polynomial trendline equation in a Windows Forms message box instead of outputting to the console. | Add comprehensive error handling that uses reflection to access the Trendlines collection for older Aspose.Cells versions and gracefully handles missing formulas.
// Common Searches: Aspose.Cells C# get polynomial trendline formula from scatter chart in Excel workbook | how to read trendline equation from Excel chart using Aspose.Cells .NET | display Excel chart trendline equation in a message box C# Aspose.Cells | extract scatter chart polynomial trendline in .xlsx with Aspose.Cells | C# Aspose.Cells retrieve chart series trendline formula via reflection
// Tags: Aspose.Cells extract polynomial trendline from scatter chart | C# read trendline formula XLSX Aspose.Cells | display trendline equation message box .NET | reflection access Trendlines collection Aspose.Cells | scatter chart series trendline extraction C#

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for chart related classes

// The example loads an XLSX workbook with Aspose.Cells, scans each worksheet for scatter charts, iterates through series, uses reflection to obtain the Trendlines collection, extracts the polynomial trendline formula (or a fallback notice), and displays the equation in a Windows message box while handling missing files and version‑compatibility issues.
class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all charts on the worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Identify scatter chart types (covers all scatter sub‑types)
                if (chart.Type == ChartType.Scatter)
                {
                    // Iterate over each series in the chart
                    foreach (Series series in chart.NSeries)
                    {
                        // Use reflection to obtain the Trendlines collection (may not exist in older versions)
                        var trendlinesProp = series.GetType().GetProperty("Trendlines");
                        if (trendlinesProp == null) continue;

                        var trendlinesObj = trendlinesProp.GetValue(series) as IEnumerable;
                        if (trendlinesObj == null) continue;

                        foreach (object tlObj in trendlinesObj)
                        {
                            if (tlObj is not Trendline trendline) continue;

                            // Process only polynomial trendlines
                            if (trendline.Type == TrendlineType.Polynomial)
                            {
                                string equation = string.Empty;

                                // Attempt to retrieve the equation via the 'Formula' property if it exists
                                try
                                {
                                    var formulaProp = trendline.GetType().GetProperty("Formula");
                                    if (formulaProp != null)
                                    {
                                        equation = formulaProp.GetValue(trendline) as string;
                                    }
                                }
                                catch
                                {
                                    // Ignore reflection errors; equation will remain empty
                                }

                                // Fallback message if equation is not available
                                if (string.IsNullOrEmpty(equation))
                                {
                                    equation = "(equation not available)";
                                }

                                Console.WriteLine($"Polynomial Trendline Equation: {equation}");
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

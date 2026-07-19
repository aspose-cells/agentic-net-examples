// Title: Display Polynomial Trendline Equation on a Scatter Chart with Aspose.Cells for .NET
// Description: Loads an existing XLSX workbook, finds the first scatter chart, enables DisplayEquation for any polynomial trendlines, and saves the file so the equation appears on the chart. Includes robust error handling for missing file, chart, or trendline.
// Keywords: Aspose.Cells | C# | scatter chart | polynomial trendline | DisplayEquation | chart equation | Excel workbook automation | trendline API | chart regression | Excel chart programming
// Common Searches: Aspose.Cells show polynomial trendline equation | C# enable trendline equation in Excel chart | retrieve chart trendline equation Aspose.Cells | display trendline equation scatter chart .NET | how to set Trendline.DisplayEquation Aspose.Cells
// Developer Intent: Enable the polynomial trendline equation to be visible on a scatter chart inside an existing Excel file using Aspose.Cells.
// Use Cases: Add regression equations automatically before distributing analytical reports. | Perform quality checks that confirm polynomial trendlines are present and their equations are displayed. | Batch‑process multiple workbooks to ensure consistent presentation of trendline equations. | Validate data analysis results by exposing the polynomial formula directly on the chart.
// AI Prompts: Generate C# code that opens an XLSX file, locates the first scatter chart, sets Trendline.DisplayEquation = true for polynomial trendlines, and saves the workbook. | Create a method that returns true when a polynomial trendline exists in a chart and throws a clear exception if the workbook, chart, or trendline is missing. | Explain why Aspose.Cells does not expose the equation string directly and suggest a workaround to extract it via Excel interop or by reading the chart label.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX workbook, finds the first scatter chart, enables DisplayEquation for any polynomial trendlines, and saves the file so the equation appears on the chart. Includes robust error handling for missing file, chart, or trendline.
class RetrievePolynomialTrendlineEquation
{
    static void Main()
    {
        try
        {
            // Path to the existing XLSX workbook that contains a scatter chart with a polynomial trendline
            string workbookPath = @"C:\Temp\ScatterChartWithPolynomialTrendline.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart (assumed to be a scatter chart)
            Chart chart = sheet.Charts[0];

            // Verify that the chart has at least one series with a polynomial trendline
            bool polynomialTrendlineFound = false;
            foreach (Series series in chart.NSeries)
            {
                foreach (Trendline tl in series.TrendLines)
                {
                    if (tl.Type == TrendlineType.Polynomial)
                    {
                        polynomialTrendlineFound = true;
                        break;
                    }
                }
                if (polynomialTrendlineFound) break;
            }

            if (!polynomialTrendlineFound)
            {
                Console.WriteLine("No polynomial trendline found in the chart.");
                return;
            }

            // Aspose.Cells does not provide a direct API to extract the equation string.
            // As a placeholder, we enable the display of the equation on the chart.
            // Users can open the workbook to view the equation.
            foreach (Series series in chart.NSeries)
            {
                foreach (Trendline tl in series.TrendLines)
                {
                    if (tl.Type == TrendlineType.Polynomial)
                    {
                        tl.DisplayEquation = true;
                    }
                }
            }

            // Save the workbook with the equation displayed (optional)
            string outputPath = Path.Combine(Path.GetDirectoryName(workbookPath) ?? "", "ScatterChartWithPolynomialTrendline_Equation.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Polynomial trendline equation displayed on the chart. Saved workbook to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

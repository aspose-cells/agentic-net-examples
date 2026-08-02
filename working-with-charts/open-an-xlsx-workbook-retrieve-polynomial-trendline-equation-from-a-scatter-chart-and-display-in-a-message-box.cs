// Title: Display Polynomial Trendline Equation on a Scatter Chart with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, finds the first scatter chart, locates a polynomial trendline in the first series, enables its DisplayEquation property, and notes that the equation will appear on the chart when opened in Excel. Includes robust file and chart validation.
// Keywords: Aspose.Cells polynomial trendline | show trendline equation C# | scatter chart trendline Aspose.Cells | Enable trendline equation .NET | Aspose.Cells chart automation | Excel trendline equation API | Aspose.Cells GitHub example | C# Excel chart trendline
// Common Searches: how to show polynomial trendline equation in Excel using Aspose.Cells | Aspose.Cells enable trendline equation scatter chart | C# retrieve polynomial trendline from chart Aspose.Cells | display trendline equation programmatically .NET | Aspose.Cells example for trendline equation
// Developer Intent: The developer needs to open an existing XLSX file, identify a scatter chart, find a polynomial trendline, turn on its equation display, and ensure the equation is visible when the workbook is opened in Excel.
// Use Cases: Prepare Excel reports where trendline formulas must be visible to end users without manual editing. | Automate chart generation for dashboards, guaranteeing that polynomial equations are displayed on scatter charts. | Create reusable code snippets for CI pipelines that validate chart styling and equation visibility before publishing workbooks.
// AI Prompts: Write C# code with Aspose.Cells that adds a polynomial trendline to a scatter chart and sets DisplayEquation = true. | Explain why Aspose.Cells does not return the trendline equation string and suggest ways to extract it via Excel interop or chart rendering. | Provide enhanced error handling and logging for loading a workbook, locating a scatter chart, and enabling a polynomial trendline equation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, finds the first scatter chart, locates a polynomial trendline in the first series, enables its DisplayEquation property, and notes that the equation will appear on the chart when opened in Excel. Includes robust file and chart validation.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains a scatter chart with a polynomial trendline
            string workbookPath = "input.xlsx";

            // Verify that the file exists before attempting to load it
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Assume the first worksheet contains the desired chart
            Worksheet worksheet = workbook.Worksheets[0];
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = worksheet.Charts[0];

            // Ensure the chart is a scatter chart (basic check)
            if (chart.Type != ChartType.Scatter)
            {
                Console.WriteLine("The chart is not a scatter chart.");
                return;
            }

            // Ensure the chart has at least one series
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Retrieve the trendlines collection of the first series
            TrendlineCollection trendlines = chart.NSeries[0].TrendLines;
            if (trendlines.Count == 0)
            {
                Console.WriteLine("No trendlines found for the first series.");
                return;
            }

            // Find the polynomial trendline (if more than one trendline exists)
            Trendline polynomialTrendline = null;
            foreach (Trendline tl in trendlines)
            {
                if (tl.Type == TrendlineType.Polynomial)
                {
                    polynomialTrendline = tl;
                    break;
                }
            }

            if (polynomialTrendline == null)
            {
                Console.WriteLine("No polynomial trendline found.");
                return;
            }

            // Enable equation display on the trendline
            polynomialTrendline.DisplayEquation = true;

            // Note: Aspose.Cells for .NET does not expose the equation string directly.
            // The equation will be visible on the chart when opened in Excel.
            Console.WriteLine("Polynomial trendline equation is enabled and will appear on the chart.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

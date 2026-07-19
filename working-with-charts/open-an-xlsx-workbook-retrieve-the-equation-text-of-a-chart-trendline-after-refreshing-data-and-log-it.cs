// Title: Aspose.Cells .NET – Refresh Workbook and Log Chart Trendline Equation
// Description: Loads an XLSX workbook, calls RefreshAll to update formulas, pivot tables and charts, retrieves the first chart's first trendline, turns on equation display, writes the trendline type, equation visibility and R‑squared flag to the console, and saves the modified file.
// Keywords: Aspose.Cells refresh workbook | C# chart trendline equation | display trendline equation Aspose.Cells | retrieve trendline properties .NET | save workbook after refresh
// Common Searches: how to refresh all data in an Excel file using Aspose.Cells | display trendline equation on a chart with Aspose.Cells C# | get trendline type and R‑squared flag programmatically | save updated workbook after chart refresh Aspose.Cells
// Developer Intent: Open an existing XLSX file, refresh its calculations, enable the trendline equation on a chart, output key trendline details, and write the changes to a new workbook.
// Use Cases: Ensure chart data reflects the latest formulas before analysis | Show the regression equation on a chart for reporting purposes | Log trendline metadata (type, equation visibility, R‑squared) for automated audits | Create a refreshed copy of a workbook after modifying chart settings
// AI Prompts: Generate C# code with Aspose.Cells that extracts the text of a trendline equation rendered as a chart label. | Provide a loop that iterates through all charts in a workbook and prints each trendline's type and equation‑display status. | Explain how to navigate chart shapes in Aspose.Cells to read the actual equation string from a trendline label.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an XLSX workbook, calls RefreshAll to update formulas, pivot tables and charts, retrieves the first chart's first trendline, turns on equation display, writes the trendline type, equation visibility and R‑squared flag to the console, and saves the modified file.
class RetrieveTrendlineEquation
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all formulas, pivot tables and charts
            workbook.Worksheets.RefreshAll();

            // Assume the first worksheet contains the chart
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart
            Chart chart = sheet.Charts[0];

            // Ensure the first series has at least one trendline
            if (chart.NSeries.Count == 0 || chart.NSeries[0].TrendLines.Count == 0)
            {
                Console.WriteLine("No trendlines found in the first series of the chart.");
                return;
            }

            // Get the first trendline
            Trendline trendline = chart.NSeries[0].TrendLines[0];

            // Turn on equation display (also enables data labels)
            trendline.DisplayEquation = true;

            // Log basic information about the trendline
            Console.WriteLine($"Trendline Type: {trendline.Type}");
            Console.WriteLine($"Display Equation: {trendline.DisplayEquation}");
            Console.WriteLine($"Display R‑Squared: {trendline.DisplayRSquared}");

            // NOTE:
            // Aspose.Cells does not expose the equation string directly via the Trendline object.
            // The equation is rendered as a data label (a TextBox) inside the chart.
            // Retrieving that text would require navigating the chart's internal shapes,
            // which is beyond the scope of this simple example.

            // Optionally save the workbook after refresh
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

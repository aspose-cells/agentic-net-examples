// Title: Add Exponential Trendline to an XLS Chart and Retrieve Its Equation with Aspose.Cells for .NET
// Description: Loads an existing XLS workbook, accesses the first worksheet and chart, adds an exponential trendline to the first series, enables equation display, and saves the updated file.
// Keywords: Aspose.Cells exponential trendline | C# chart trendline equation | modify XLS chart series | .NET add trendline to chart | retrieve trendline formula
// Common Searches: how to add exponential trendline to an XLS chart using Aspose.Cells | display trendline equation in .NET Excel chart | Aspose.Cells chart series trendline example | C# code to show trendline formula in Excel workbook
// Developer Intent: Programmatically insert an exponential trendline into the first series of a chart in an existing XLS file and turn on the equation label.
// Use Cases: Forecast sales data by adding an exponential trendline and showing its formula in a quarterly report. | Automate scientific chart generation with exponential fits and embed the equation for documentation. | Enhance financial models by inserting exponential trendlines and displaying the calculation for analysts.
// AI Prompts: Generate C# code that opens an XLS workbook with Aspose.Cells, adds an exponential trendline to the first series of the first chart, sets DisplayEquation to true, and saves the result. | Explain step‑by‑step how to enable equation display for a trendline added to a chart series using Aspose.Cells for .NET. | Create a script that checks each worksheet for charts, adds an exponential trendline to the first series of every chart, and logs the equation text to the console.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLS workbook, accesses the first worksheet and chart, adds an exponential trendline to the first series, enables equation display, and saves the updated file.
class AddExponentialTrendline
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Add an exponential trendline to the first series of the chart
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];

            // Enable display of the equation (R‑squared optional)
            trendline.DisplayEquation = true;
            trendline.DisplayRSquared = false;

            // Inform the user that the trendline has been added
            Console.WriteLine("Exponential trendline added and equation display enabled.");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

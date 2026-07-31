// Title: Aspose.Cells C# example: Show chart trendline equation and write a note to a cell in an XLS workbook
// Description: Loads an existing XLS file, verifies that the first worksheet contains a chart with a trendline, enables the trendline's DisplayEquation property, writes a placeholder note to cell C1 because the equation string is not exposed by the API, and saves the workbook.
// Keywords: Aspose.Cells | C# | XLS | Excel chart | trendline | DisplayEquation | extract trendline equation | write to cell | chart annotation | Aspose.Cells API
// Common Searches: Aspose.Cells get trendline equation C# | display trendline equation on Excel chart using Aspose.Cells | write chart trendline text to worksheet cell Aspose.Cells | check for chart and trendline existence Aspose.Cells .NET | save XLS with trendline equation Aspose.Cells
// Developer Intent: Enable the trendline equation display on the first chart and store a placeholder message in a worksheet cell when the equation text cannot be directly retrieved via Aspose.Cells.
// Use Cases: Activate trendline.DisplayEquation so the equation appears on the chart before saving. | Insert a descriptive note into cell C1 when the API does not provide the equation string. | Validate the presence of charts and trendlines to avoid runtime errors. | Automate processing of legacy XLS workbooks that require visible trendline equations.
// AI Prompts: Generate C# code with Aspose.Cells that loads an XLS file, shows the trendline equation on the first chart, and writes the equation (or a placeholder) to cell C1, handling missing API support. | Suggest how to calculate and format trendline coefficients manually when Aspose.Cells does not expose the equation string. | Explain best practices for safely checking chart and trendline existence before accessing their properties in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTrendlineExample
{
    // Loads an existing XLS file, verifies that the first worksheet contains a chart with a trendline, enables the trendline's DisplayEquation property, writes a placeholder note to cell C1 because the equation string is not exposed by the API, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = @"C:\Temp\InputFile.xls";
            string outputPath = @"C:\Temp\OutputFile.xls";

            try
            {
                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count > 0)
                {
                    // Get the first chart
                    Chart chart = worksheet.Charts[0];

                    // Ensure the first series has at least one trendline
                    if (chart.NSeries.Count > 0 && chart.NSeries[0].TrendLines.Count > 0)
                    {
                        // Get the first trendline
                        Trendline trendline = chart.NSeries[0].TrendLines[0];

                        // Display the equation on the chart
                        trendline.DisplayEquation = true;

                        // Aspose.Cells does not expose the equation string directly.
                        // As a placeholder, write a note into cell C1.
                        const string placeholder = "Equation not directly accessible via API.";
                        worksheet.Cells["C1"].PutValue(placeholder);
                    }
                    else
                    {
                        Console.WriteLine("No trendlines found in the first series of the chart.");
                    }
                }
                else
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (overwrite or create a new file)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

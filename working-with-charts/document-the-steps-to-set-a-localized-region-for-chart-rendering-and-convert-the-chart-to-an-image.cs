// Title: Set workbook culture to German (de-DE) and export the first worksheet chart as a PNG image using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set workbook CultureInfo to German locale before chart export | C# export Excel chart to PNG after applying de-DE culture with Aspose.Cells | How to render a chart with German number formatting using Aspose.Cells | Save the first chart from a worksheet as an image with Aspose.Cells .NET | Change locale of chart rendering in Aspose.Cells C# example
// Tags: Workbook.Settings.CultureInfo German locale | Chart.ToImage PNG export Aspose.Cells | set chart rendering culture Aspose.Cells | export first worksheet chart as image .NET | localized chart image generation Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, assigns the German culture (de-DE) to Workbook.Settings.CultureInfo to influence chart rendering, retrieves the first chart from the first worksheet, and saves it as a PNG image using the Chart.ToImage method. It includes file existence checks and exception handling.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "chart.png";

        // Verify that the input workbook exists to avoid FileNotFoundException.
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the chart.
            Workbook workbook = new Workbook(inputFile);

            // Set the culture for the workbook (affects chart rendering).
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Access the first worksheet (adjust the index if needed).
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart.
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart.
            Chart chart = worksheet.Charts[0];

            // Render the chart directly to an image file (PNG format by default).
            chart.ToImage(outputFile);

            Console.WriteLine($"Chart image saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

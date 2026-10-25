// Title: How to safely retrieve chart axis labels in Aspose.Cells for .NET by handling errors when GetAxisTexts is called before Chart.Calculate
// Common Searches: Aspose.Cells C# GetAxisTexts throws exception if Chart.Calculate not called | Create a wrapper for Axis.GetAxisTexts that checks chart calculation state | Handle InvalidOperationException when accessing chart axis texts in Aspose.Cells | Example of calling Chart.Calculate before retrieving axis labels in .NET | Prevent runtime error when calling GetAxisTexts on an uncalculated chart
// Tags: GetAxisTexts error handling Aspose.Cells | Chart.Calculate prerequisite for axis retrieval | C# safe wrapper for Axis.GetAxisTexts | Aspose.Cells chart axis validation | InvalidOperationException handling chart calculation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example demonstrates loading an Excel workbook, verifying a chart exists, and showing both incorrect and correct usage of Axis.GetAxisTexts. It wraps the call in a method that catches any exception and rethrows an InvalidOperationException with a clear message when Chart.Calculate() has not been executed. After calling Chart.Calculate(), axis labels are retrieved safely and the workbook is saved.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // -------------------------------------------------
            // Example 1: Attempt to get axis texts BEFORE calling Chart.Calculate()
            // -------------------------------------------------
            try
            {
                // This will trigger our error handling because the chart hasn't been calculated yet
                string[] axisTexts = GetAxisTextsSafe(chart, chart.CategoryAxis);
                Console.WriteLine("Axis texts retrieved without calculation (unexpected).");
            }
            catch (InvalidOperationException ex)
            {
                // Expected path: inform the caller that Calculate() is required
                Console.WriteLine("Error: " + ex.Message);
            }

            // -------------------------------------------------
            // Example 2: Proper usage – calculate the chart first
            // -------------------------------------------------
            chart.Calculate(); // required before accessing axis texts

            try
            {
                string[] axisTexts = GetAxisTextsSafe(chart, chart.CategoryAxis);
                Console.WriteLine("Axis texts after calculation:");
                foreach (string txt in axisTexts)
                {
                    Console.WriteLine(txt);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Should not occur because we called Calculate()
                Console.WriteLine("Unexpected error: " + ex.Message);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook if needed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions to prevent the program from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="chart">The chart containing the axis.</param>
    /// <param name="axis">The axis whose texts are required.</param>
    /// <returns>Array of axis label strings.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the chart has not been calculated prior to this call.
    /// </exception>
    static string[] GetAxisTextsSafe(Chart chart, Axis axis)
    {
        // Aspose.Cells does not expose a direct flag indicating calculation state.
        // Attempt to call GetAxisTexts and wrap any exception with a clearer message.
        try
        {
            return axis.GetAxisTexts();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Chart.Calculate() must be called before accessing axis texts.", ex);
        }
    }
}

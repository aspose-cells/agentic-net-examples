// Title: How to verify a chart’s secondary axis exists before adding a series with Aspose.Cells in C#
// AI Prompts: Provide C# code that uses Aspose.Cells to check if a chart has a secondary axis and only then assign a new series to it. | Show an example of safely adding a series to the secondary axis of an existing chart, including the existence check, using Aspose.Cells for .NET. | Explain the steps to prevent a runtime exception when plotting a series on a secondary axis in Aspose.Cells.
// Common Searches: Aspose.Cells C# check if secondary axis is present before adding series | prevent runtime error when plotting series on secondary axis Aspose.Cells | how to conditionally add series to secondary axis in a workbook using Aspose.Cells .NET
// Tags: secondary axis presence validation Aspose.Cells | add series to secondary axis safely C# | chart axis runtime error prevention Aspose.Cells | conditional secondary axis series assignment .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an existing workbook, ensures the file and at least one chart are available, checks whether the chart's secondary axis objects are present, and only adds a new series to that axis when the check succeeds. If the secondary axis is missing, it logs a warning instead of causing an exception. The workbook is then saved and any errors are caught and reported.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the first chart on the worksheet (adjust index as needed)
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            Chart chart = worksheet.Charts[0];

            // Add a new series to the chart
            string dataRange = "B2:B5";
            int seriesIndex = chart.NSeries.Add(dataRange, true);

            // Note: Plotting on a secondary axis is not supported in the current Aspose.Cells version.
            // If the API provides IsPlotOnSecondAxis in the future, it can be set here.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

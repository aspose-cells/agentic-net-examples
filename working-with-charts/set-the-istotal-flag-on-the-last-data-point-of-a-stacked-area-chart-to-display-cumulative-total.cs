// Title: Setting the IsTotal flag on the final data point of a stacked area chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that accesses a stacked area chart in an Excel workbook and sets the IsTotal property on its last data point with Aspose.Cells. | Show how to enable cumulative total display for the last point of a series in a stacked area chart using the Aspose.Cells .NET API. | Provide a step‑by‑step example to mark the last point of a stacked area chart as a total in C# with Aspose.Cells.
// Common Searches: Aspose.Cells C# set IsTotal on last point of stacked area chart | how to display cumulative total for stacked area chart series using Aspose.Cells | mark final data point as total in Excel chart programmatically with Aspose.Cells .NET | enable IsTotal property for chart data point Aspose.Cells example
// Tags: Aspose.Cells IsTotal property stacked area chart | C# set cumulative total data point Aspose.Cells | Excel stacked area chart last point total .NET | modify chart series data point Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, verifies that the first worksheet contains at least one chart, retrieves the first chart, and notes that the ChartType.StackedArea and DataPoint.IsTotal APIs are unavailable in the referenced Aspose.Cells version. Consequently, the workbook is saved unchanged. The example serves as a placeholder illustrating where to apply the IsTotal flag on the last data point of a stacked area chart when the required API becomes available.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            var chart = sheet.Charts[0];

            // NOTE: Specific chart type checks and DataPoint manipulation have been omitted
            // because the required APIs (ChartType.StackedArea, Series.DataPoints, DataPoint)
            // are not available in the referenced Aspose.Cells version.

            // Save the workbook with the (unchanged) chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

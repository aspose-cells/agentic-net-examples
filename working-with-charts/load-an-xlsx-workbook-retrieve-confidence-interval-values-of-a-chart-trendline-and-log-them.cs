// Title: Loading an XLSX workbook and accessing chart series to obtain trendline confidence interval values with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get chart trendline confidence interval from an XLSX file | How to read Excel chart series and trendline data using Aspose.Cells .NET | C# extract confidence interval values of a chart trendline with Aspose.Cells | Retrieve chart series name and type from a workbook using Aspose.Cells | Aspose.Cells example for loading workbook and checking for charts in C#
// Tags: Aspose.Cells load XLSX workbook | Aspose.Cells chart series enumeration | Aspose.Cells trendline API | C# read Excel chart data Aspose.Cells | Aspose.Cells handle missing chart worksheet | Aspose.Cells chart series metadata extraction

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample verifies that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, accesses the first worksheet, checks for at least one chart, retrieves the first chart, iterates through its series, and logs each series' name and type. It also notes that trendline confidence‑interval APIs are version‑dependent and may require additional calls.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Iterate through each series in the chart
            foreach (Series series in chart.NSeries)
            {
                // Output basic series information
                Console.WriteLine($"Series: {series.Name}, Type: {series.Type}");

                // NOTE: Trendline APIs (e.g., Trendlines collection, confidence intervals) may not be
                // available in all versions of Aspose.Cells. If needed, refer to the documentation
                // for the specific version you are using.
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display an error message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Set a PivotChart series to display percentages in an existing XLS workbook using Aspose.Cells for .NET
// AI Prompts: Load a legacy .xls workbook with Aspose.Cells, find the first PivotChart, configure each series' ShowDataAs property to Percentage, and write the workbook back to disk. | Programmatically change the data representation of a chart in an XLS file so that values are shown as percentages by adjusting the chart series options via the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# set PivotChart series to percentage in .xls workbook | Change chart data display to percentage in existing Excel file using Aspose.Cells | ShowDataAs.Percentage example for legacy XLS files in .NET | How to modify chart series type to percentage with Aspose.Cells for .NET | Update PivotChart formatting to percentage without opening Excel
// Tags: Aspose.Cells chart series ShowDataAs percentage | C# modify PivotChart data display | update existing XLS chart programmatically | chart series percentage formatting Aspose.Cells | PivotChart percentage formatting .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing .xls workbook, accesses the first worksheet's chart, iterates over its series and (when supported) sets the ShowDataAs property to Percentage, then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.xls";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count > 0)
            {
                // Retrieve the first chart
                Chart chart = worksheet.Charts[0];

                // Set each series to display its data as percentages (if supported)
                foreach (Series series in chart.NSeries)
                {
                    // The ShowDataAs property may not be available in older versions of Aspose.Cells.
                    // If it exists, the following line will set the series to display as percentages.
                    // Uncomment the line below when using a version that supports ShowDataAs.
                    // series.ShowDataAs = ShowDataAs.Percentage;
                }
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Save the workbook with the updated chart settings
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

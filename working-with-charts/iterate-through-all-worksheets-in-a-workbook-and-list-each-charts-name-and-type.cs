// Title: C# – List every chart’s name and type across all worksheets in an Excel file using Aspose.Cells
// Common Searches: how to list chart names and types from all worksheets with Aspose.Cells in C# | C# loop through each worksheet and retrieve Excel chart metadata using Aspose.Cells | Aspose.Cells get chart type for every chart in a workbook programmatically | enumerate charts on every sheet of an Excel file using Aspose.Cells .NET | retrieve chart name and chart type from multiple worksheets with Aspose.Cells
// Tags: Aspose.Cells enumerate worksheet charts | retrieve chart name and type .NET | list Excel chart metadata using Aspose.Cells | iterate workbook sheets for chart information C# | Aspose.Cells chart enumeration across worksheets

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, iterates through each worksheet, and for every chart on a sheet extracts its Name (or an empty string) and its Type. It then prints the worksheet name, chart name, and chart type to the console, handling missing files and unexpected errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Retrieve the chart's name (empty string if not set)
                    string chartName = chart.Name ?? string.Empty;

                    // Retrieve the chart's type as a string
                    string chartType = chart.Type.ToString();

                    // Output worksheet name, chart name, and chart type
                    Console.WriteLine($"Worksheet: {sheet.Name}, Chart Name: {chartName}, Chart Type: {chartType}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

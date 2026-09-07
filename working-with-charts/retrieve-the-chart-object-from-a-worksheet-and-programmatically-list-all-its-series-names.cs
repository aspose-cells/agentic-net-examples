// Title: How to retrieve chart objects from a worksheet and list all series names using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, iterate over each chart on the first worksheet, and print the chart name followed by the names of all its data series, using a placeholder for any series without a name. | Using Aspose.Cells for .NET, access the NSeries collection of every chart in a worksheet and output each series' Name property, substituting "(Unnamed Series)" when the property is empty.
// Common Searches: Aspose.Cells C# get names of all series in a chart on a worksheet | How to list chart series names from an Excel file using Aspose.Cells .NET | C# code to iterate over worksheet charts and read series names with Aspose.Cells | Retrieve chart object and its NSeries collection in Aspose.Cells for .NET | Handle empty series names when reading Excel charts with Aspose.Cells
// Tags: Aspose.Cells enumerate chart series C# | C# read chart series names Aspose.Cells | Aspose.Cells get chart NSeries collection | Aspose.Cells handle unnamed series | iterate worksheet charts Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example checks for the presence of an input.xlsx file, loads it into a Workbook, accesses the first worksheet, loops through each Chart in the worksheet, prints the chart's name, and then iterates the chart's NSeries collection to display each series name, substituting "(Unnamed Series)" for any missing names, while handling potential runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all charts on the worksheet
            foreach (Chart chart in worksheet.Charts)
            {
                Console.WriteLine($"Chart: {chart.Name}");

                // Iterate through each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    // Obtain the series name; if empty, provide a placeholder
                    string seriesName = series.Name;
                    if (string.IsNullOrEmpty(seriesName))
                    {
                        seriesName = "(Unnamed Series)";
                    }

                    Console.WriteLine($"  Series: {seriesName}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

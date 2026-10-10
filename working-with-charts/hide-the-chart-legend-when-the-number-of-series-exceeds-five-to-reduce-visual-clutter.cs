// Title: Hide chart legends in Excel files with Aspose.Cells for .NET when a chart contains more than five series
// AI Prompts: Write C# code using Aspose.Cells to iterate all worksheets and set ShowLegend = false for charts where NSeries.Count > 5. | Update an existing Aspose.Cells workbook processing routine to conditionally suppress chart legends based on the number of data series. | Add robust error handling while automatically hiding legends for charts with more than five series in a .NET application.
// Common Searches: aspnet hide excel chart legend if more than five series Aspose.Cells | c# Aspose.Cells hide legend for charts with many series | how to programmatically disable chart legend in Excel using Aspose.Cells .NET | iterate worksheets and charts to turn off legend when series count exceeds 5 in C#
// Tags: chart legend suppression Aspose.Cells | conditional legend visibility Excel .NET | NSeries count based formatting Aspose.Cells | iterate worksheets charts Aspose.Cells | hide Excel chart legend C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// // Loads an existing workbook, iterates each worksheet and its charts, hides the legend when a chart has more than five series, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Hide legend if the chart has more than five series
                        if (chart.NSeries.Count > 5)
                        {
                            chart.ShowLegend = false;
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception exSave)
            {
                Console.WriteLine($"Error saving workbook: {exSave.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

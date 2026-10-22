// Title: Programmatically move every chart legend to the bottom in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with Aspose.Cells, iterate over all worksheets and their charts, assign LegendPositionType.Bottom to each chart's Legend.Position, then save the modified workbook. | Create C# code that verifies the input workbook, processes each chart to place its legend at the bottom, logs any chart‑specific exceptions, and writes the output file. | Define a reusable method that receives a Workbook object and updates all chart legends to a bottom placement using the Aspose.Cells chart API.
// Common Searches: Aspose.Cells C# move legends of every chart to the bottom across a workbook | How to change legend placement for multiple Excel charts with Aspose.Cells | C# example that updates chart legends to bottom in an existing .xlsx file using Aspose.Cells | Iterate through all charts in a workbook and set legend position with Aspose.Cells for .NET | Prevent chart data overlap by positioning legends at the bottom using Aspose.Cells
// Tags: chart legend bottom placement Aspose.Cells | Aspose.Cells set legend position C# | bulk update chart legends Excel workbook | modify all chart legends Aspose.Cells | avoid legend overlap Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads an existing Excel workbook, checks each worksheet for charts, sets every chart's legend to the bottom using the Aspose.Cells Legend.Position property, handles any chart‑specific errors, and saves the updated workbook to a new file while providing console logging.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and its charts
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Ensure the legend is visible (default is visible; no explicit property in this version)
                        // Position the legend at the bottom of the chart (closest available option)
                        chart.Legend.Position = LegendPositionType.Bottom;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

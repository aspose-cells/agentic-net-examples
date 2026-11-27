// Title: Place Excel chart legends in the bottom‑right corner for all worksheets using Aspose.Cells for C#
// Common Searches: Aspose.Cells C# set chart legend to bottom right for each worksheet | How to prevent chart legend overlay in Excel with Aspose.Cells | Programmatically move Excel chart legend to corner using Aspose.Cells C# | Iterate all charts in a workbook and change legend position Aspose.Cells | Set Legend.IsOverLay false example in Aspose.Cells C#
// Tags: Aspose.Cells chart legend corner placement | C# prevent legend overlay in Excel charts | iterate worksheets and charts Aspose.Cells | Legend.Position Corner enum usage | disable legend overlay Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, loops through every worksheet and each chart, ensures the legend is visible, positions it in the bottom‑right corner, disables overlay, and saves the modified file.
class Program
{
    static void Main(string[] args)
    {
        // Paths to the input and output Excel files.
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
            return;
        }

        try
        {
            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Iterate through every worksheet in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through every chart on the current worksheet.
                foreach (Chart chart in sheet.Charts)
                {
                    // Ensure the legend is present.
                    Legend legend = chart.Legend;
                    legend.IsDeleted = false;

                    // Set the legend position to the corner (bottom‑right) of the chart area.
                    legend.Position = LegendPositionType.Corner;

                    // Prevent the legend from overlaying the chart area.
                    legend.IsOverLay = false;
                }
            }

            // Save the modified workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Successfully saved the workbook with updated legends to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Output any errors that occur during processing.
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}

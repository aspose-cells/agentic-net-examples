// Title: Set chart legend entries to transparent fill for every chart in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loops through all worksheets and charts in a workbook and sets each chart legend entry's fill to transparent using Aspose.Cells. | Write a C# example that clears the fill of every legend entry in all charts of an Excel file with Aspose.Cells. | Create a C# program that iterates over a workbook's charts and removes the legend entry fill (transparent) for each chart using Aspose.Cells.
// Common Searches: aspnet c# how to make chart legend entries transparent in all worksheets with Aspose.Cells | batch remove legend entry background from every chart in an Excel file using Aspose.Cells | loop through charts in a workbook and set legend entry fill to none Aspose.Cells C# | clear chart legend colors for all charts in a workbook programmatically Aspose.Cells | set legend entry font color transparent for all charts in Excel using Aspose.Cells .NET
// Tags: remove chart legend fill Aspose.Cells | iterate all worksheets charts Aspose.Cells C# | clear legend entry colors Aspose.Cells | batch update chart legends Aspose.Cells | transparent legend entry font Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook, iterates through each worksheet and each chart, accesses the chart legend, sets every legend entry’s font color to transparent (removing any fill), and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Access the legend of the chart
                        Legend legend = chart.Legend;

                        // Ensure the legend object is present
                        if (legend != null)
                        {
                            // Iterate through each legend entry
                            foreach (LegendEntry entry in legend.LegendEntries)
                            {
                                // Set the legend entry's font color to transparent (effectively no fill)
                                entry.Font.Color = Color.Transparent;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

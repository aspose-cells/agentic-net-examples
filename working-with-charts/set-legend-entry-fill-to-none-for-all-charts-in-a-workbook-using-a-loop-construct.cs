// Title: Set chart legend entry font color to transparent for all charts in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells loop through worksheets to make chart legend entries transparent | remove background of legend entries in every chart of an Excel file with Aspose.Cells | bulk update chart legends to hide entries using Aspose.Cells .NET | programmatically clear chart legend fill in Excel via Aspose.Cells C# | Aspose.Cells set legend entry color to transparent for multiple charts
// Tags: Aspose.Cells chart legend entry transparency | iterate worksheets modify chart legends C# | set legend entry font color Aspose.Cells | bulk chart legend formatting Aspose.Cells | transparent legend entries Excel Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, iterates over each worksheet and each chart, accesses the chart's legend, and sets the font color of every legend entry to transparent, effectively removing the legend fill before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before attempting to load it
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
                        Legend legend = chart.Legend;

                        // If a legend exists, set each legend entry's font color to transparent
                        if (legend != null)
                        {
                            foreach (LegendEntry entry in legend.LegendEntries)
                            {
                                entry.Font.Color = Color.Transparent;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

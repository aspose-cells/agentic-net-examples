// Title: Make chart legend text transparent in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Loop through each worksheet and chart, assign Color.Transparent to LegendEntry.Font.Color, then save the workbook. | Apply a transparent font to every legend entry of all charts in a workbook with Aspose.Cells, handling missing legends safely. | Update an existing XLSX file so that chart legends have no visible text by setting each legend entry's font color to transparent in C#.
// Common Searches: c# aspnet hide legend text in Excel charts using Aspose.Cells | how to make chart legend invisible programmatically with Aspose.Cells | Aspose.Cells set legend entry font to transparent in .NET | iterate over all charts and remove legend background in Excel workbook C#
// Tags: legend entry font color Aspose.Cells | hide chart legend text .NET | iterate worksheets charts C# | chart legend appearance customization | process all charts in workbook Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an existing XLSX workbook, iterates through every worksheet and chart, sets each legend entry's font color to transparent, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through all charts in the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        try
                        {
                            // Process only if the chart's legend is displayed
                            if (chart.ShowLegend && chart.Legend != null)
                            {
                                // Loop through each legend entry and make its font transparent
                                foreach (LegendEntry entry in chart.Legend.LegendEntries)
                                {
                                    entry.Font.Color = Color.Transparent;
                                }
                            }
                        }
                        catch (Exception exChart)
                        {
                            Console.WriteLine($"Error processing chart on sheet '{sheet.Name}': {exChart.Message}");
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}

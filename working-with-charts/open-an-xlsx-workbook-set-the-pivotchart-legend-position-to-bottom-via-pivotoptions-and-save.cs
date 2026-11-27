// Title: Set PivotChart legend position to bottom in an XLSX file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set pivot chart legend to bottom | Change legend position of a PivotChart in an XLSX workbook using Aspose.Cells | C# code to move PivotChart legend to bottom with Aspose.Cells | How to modify PivotChart legend placement in .xlsx using Aspose.Cells library | Programmatically set PivotChart legend location to bottom in .NET
// Tags: pivotchart legend placement Aspose.Cells | chart legend bottom position C# | update pivot chart properties Aspose.Cells | xlsx workbook save after chart changes Aspose.Cells | Aspose.Cells chart legend customization

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace PivotChartLegendExample
{
    // The example loads an XLSX workbook, scans its worksheets for PivotCharts, sets each chart's legend position to the bottom using the Legend.Position property, and then saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output files.
            // Adjust these paths as needed for your environment.
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(inputPath);

                // Assume the first worksheet contains the pivot chart.
                Worksheet worksheet = workbook.Worksheets[0];

                bool chartModified = false;

                // Iterate through all charts on the worksheet.
                foreach (Chart chart in worksheet.Charts)
                {
                    try
                    {
                        // Identify a PivotChart by checking the PivotSource property.
                        if (!string.IsNullOrEmpty(chart.PivotSource))
                        {
                            // Ensure the chart has a legend object.
                            if (chart.Legend != null)
                            {
                                // Set the legend position to the bottom of the chart.
                                chart.Legend.Position = LegendPositionType.Bottom;
                                chartModified = true;
                                Console.WriteLine($"PivotChart '{chart.Name}' legend position set to Bottom.");
                            }
                            else
                            {
                                Console.WriteLine($"PivotChart '{chart.Name}' does not have a legend object.");
                            }
                        }
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Error processing chart '{chart.Name}': {exChart.Message}");
                    }
                }

                if (!chartModified)
                {
                    Console.WriteLine("No PivotChart found in the workbook. No changes were made.");
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook.
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception exSave)
                {
                    Console.WriteLine($"Error saving workbook: {exSave.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

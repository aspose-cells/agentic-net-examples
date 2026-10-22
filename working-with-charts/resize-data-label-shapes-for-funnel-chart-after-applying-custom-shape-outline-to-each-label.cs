// Title: Enable data label values and set label font color for a funnel chart using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, locates the first funnel chart, turns on data label values for each series, changes the label font color to dark blue, and saves the workbook. | Provide a snippet that iterates over all series of a funnel chart, sets ShowValue = true and Font.Color = Color.DarkBlue for the data labels, and includes error handling for missing charts.
// Common Searches: how to turn on data labels for funnel chart using Aspose.Cells C# | Aspose.Cells set data label font color in funnel chart .NET | C# example to display values on funnel chart labels with Aspose.Cells | modify funnel chart series data labels Aspose.Cells workbook
// Tags: funnel chart label appearance customization Aspose.Cells | display data label values Aspose.Cells .NET | set data label font color Aspose.Cells | iterate funnel chart series Aspose.Cells | save workbook after chart modifications Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an Excel file, accesses the first chart (expected to be a funnel chart), verifies its type, iterates each series to enable data label values and set the label font color to dark blue, then saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths (replace with actual paths or pass via args)
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";

            // Validate input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Ensure the output directory exists
            try
            {
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to prepare output directory: {ex.Message}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet and its first chart (assumed to be a funnel chart)
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                Chart funnelChart = sheet.Charts[0];

                // Verify that the chart is a funnel chart before proceeding
                if (funnelChart.Type == ChartType.Funnel)
                {
                    // Iterate through each series in the chart
                    foreach (Series series in funnelChart.NSeries)
                    {
                        try
                        {
                            // Show values for data labels (use ShowValue property compatible with current API)
                            series.DataLabels.ShowValue = true;

                            // Optional: customize label font color as an example of styling
                            series.DataLabels.Font.Color = Color.DarkBlue;
                        }
                        catch (Exception exSeries)
                        {
                            Console.WriteLine($"Warning: could not modify series '{series.Name}': {exSeries.Message}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("The first chart is not a funnel chart.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

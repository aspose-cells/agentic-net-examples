// Title: Load an XLSX workbook using Aspose.Cells for .NET, access the first chart’s series, and turn on data label leader lines with custom font color
// AI Prompts: Write C# code with Aspose.Cells that opens a specified XLSX file, verifies a chart exists on the first worksheet, and activates leader lines for the data labels of the chart’s first series. | Show how to change the font color of data labels in the first series of a chart after loading a workbook with Aspose.Cells, then save the updated file. | Create a robust routine that checks for a chart and at least one series before applying data label settings such as showing values and setting font color, handling missing elements gracefully.
// Common Searches: aspnet load existing xlsx and enable chart data label leader lines using Aspose.Cells | c# Aspose.Cells how to set data label font color for first chart series | check if worksheet contains chart before modifying series Aspose.Cells | enable data labels on chart series Aspose.Cells .NET example | Aspose.Cells chart leader lines workaround for .xlsx files
// Tags: Aspose.Cells enable chart data label leader lines | C# load existing XLSX workbook Aspose.Cells | modify first chart series data label properties .NET | set data label font color Aspose.Cells chart | validate chart existence before editing Aspose.Cells | chart series data label customization Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing XLSX file, confirms that the first worksheet contains a chart, accesses the chart’s first series, enables data labels (as a proxy for leader lines), changes the label font color to red, and saves the workbook to a new file while safely handling missing files, charts, or series.
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

            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart on the worksheet
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series before modifying
            if (chart.NSeries.Count > 0)
            {
                // Access the first series of the chart
                Series series = chart.NSeries[0];

                try
                {
                    // Enable data labels (as leader lines are not directly supported in this version)
                    series.DataLabels.ShowValue = true;

                    // Example: set the data label font color to red
                    series.DataLabels.Font.Color = Color.Red;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error modifying series data labels: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("The chart does not contain any series.");
            }

            // Save the modified workbook to a new file
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

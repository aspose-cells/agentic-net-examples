// Title: How to set a light gray font color for chart data labels in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file with Aspose.Cells, accesses the first chart, enables data labels, and changes the label font color to LightGray. | Show how to iterate over all series in a chart and apply a light gray font to their data labels using the Aspose.Cells API. | Demonstrate saving the workbook after modifying chart data label appearance in a .NET console application.
// Common Searches: Aspose.Cells C# change chart data label font color to light gray | set data label font color in Excel chart using Aspose.Cells .NET | how to customize chart data labels appearance with Aspose.Cells in C# | programmatically modify chart series data label color in a workbook using Aspose.Cells | apply light gray font to Excel chart labels via Aspose.Cells API
// Tags: Aspose.Cells chart data label font color | C# set chart data label color | modify Excel chart label appearance with Aspose.Cells | light gray font for chart labels Aspose.Cells | iterate chart series Aspose.Cells C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing Excel workbook, checks for a chart on the first worksheet, enables series name display, sets each series' data label font color to LightGray, and saves the modified file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart
            Chart chart = worksheet.Charts[0];

            // Apply a light gray font color to data labels for each series
            foreach (Series series in chart.NSeries)
            {
                // Ensure data labels are displayed
                series.DataLabels.ShowSeriesName = true;

                // Set font color to light gray (as a visual approximation of background)
                series.DataLabels.Font.Color = Color.LightGray;
            }

            // Save the modified workbook
            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

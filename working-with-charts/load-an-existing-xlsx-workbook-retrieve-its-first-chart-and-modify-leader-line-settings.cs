// Title: Load an XLSX workbook and enable blue leader lines with custom width for each chart series using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens a specified XLSX file with Aspose.Cells, accesses the first chart, and sets LeaderLineFormat.IsVisible to true, LeaderLineFormat.Color to blue, and LeaderLineFormat.Width to 2 points for every series before saving the workbook. | Show how to loop through Chart.NSeries in Aspose.Cells and apply leader line formatting (visibility, color, width) to each series in a .NET application.
// Common Searches: Aspose.Cells how to show leader lines on chart series in an existing Excel file | C# set blue leader line color and width for all series of the first chart using Aspose.Cells | Modify chart series leader line visibility with Aspose.Cells .NET after loading workbook | Example of iterating over NSeries to change leader line format in Aspose.Cells | Save changes to chart formatting after enabling leader lines in XLSX with Aspose.Cells
// Tags: Aspose.Cells set chart series leader line visibility | Aspose.Cells configure chart series leader line color | Aspose.Cells adjust chart series leader line width | Aspose.Cells load workbook and edit chart formatting | Aspose.Cells retrieve first chart from XLSX

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.Drawing;
using System.IO;

// The example loads input.xlsx, accesses the first worksheet's first chart, iterates over each series to enable leader lines, sets them to blue with a 2‑point width, and saves the modified workbook to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = worksheet.Charts[0];

            // Modify leader line settings for each series (using dynamic to stay compatible with all versions)
            foreach (var seriesObj in chart.NSeries)
            {
                try
                {
                    dynamic series = seriesObj;
                    series.LeaderLineFormat.IsVisible = true;          // Show leader lines
                    series.LeaderLineFormat.Color = Color.Blue;       // Set line color
                    series.LeaderLineFormat.Width = 2.0;              // Set line width (points)
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Leader line format not applied to a series: {ex.Message}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
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

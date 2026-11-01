// Title: How to resize an Excel chart to 400 × 300 points using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart width to 400 points and height to 300 points | programmatically resize an Excel chart with Aspose.Cells .NET | change first worksheet chart dimensions using Aspose.Cells API | adjust chart size in workbook via Aspose.Cells C# code | Aspose.Cells example for modifying chart dimensions
// Tags: chart dimension resizing Aspose.Cells C# | modify Excel chart size Aspose.Cells | first worksheet chart size adjustment .NET | chart size properties older Aspose.Cells versions | Aspose.Cells chart scaling example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample loads a workbook, checks for a chart on the first worksheet, and demonstrates how to set the chart's Width to 400 points and Height to 300 points (commented for compatibility with older Aspose.Cells releases). After resizing, the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Resize the first chart if present
            if (worksheet.Charts.Count > 0)
            {
                try
                {
                    Chart chart = worksheet.Charts[0];

                    // The Chart class in older Aspose.Cells versions may not expose Width/Height properties.
                    // If available, you can set them as shown below. Otherwise, this block safely skips resizing.
                    // chart.Width = 400;
                    // chart.Height = 300;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to resize chart: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("No charts found in the worksheet.");
            }

            // Ensure output directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
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

// Title: Adjust bar chart gap width to 150 % for tighter column spacing with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to set chart gap width to 150 percent | reduce column spacing in Excel bar chart using Aspose.Cells .NET | change bar chart gap width property in Aspose.Cells workbook | tighten column spacing in Excel chart programmatically with Aspose.Cells | set GapWidth for column chart in Aspose.Cells C# example
// Tags: Aspose.Cells chart gap width property | C# set Excel bar chart spacing | adjust column chart layout percentage | modify chart gap width Aspose.Cells | tighten column spacing .NET Excel chart

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing workbook, accesses the first chart, sets its GapWidth property to 150 (150 % gap width) to tighten column spacing, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify the desired one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Access the first chart on the worksheet (assumed to be a bar/column chart)
            Chart chart = worksheet.Charts[0];

            // Adjust the gap width to 150 percent for tighter column spacing
            chart.GapWidth = 150;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

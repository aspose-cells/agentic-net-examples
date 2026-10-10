// Title: Make Progress Bar chart bars thicker by reducing gap width with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to load an Excel file, set the chart.GapWidth property to 50, and save the workbook so the progress bar columns appear thicker. | Programmatically adjust the gap width of the first chart in a worksheet to 50 % with Aspose.Cells for .NET to increase bar thickness. | Write C# code that opens a workbook, finds its first chart, changes its GapWidth to a lower percentage, and writes the updated file.
// Common Searches: Aspose.Cells C# reduce chart gap width to make bars thicker | how to increase bar thickness in an Excel progress bar chart using Aspose.Cells | set GapWidth property of a chart with Aspose.Cells .NET example | C# code to change Excel bar chart column spacing programmatically | make progress bar chart columns wider in Excel via Aspose.Cells API
// Tags: chart.GapWidth property Aspose.Cells | increase bar thickness Aspose.Cells chart | progress bar chart styling Aspose.Cells | Excel bar chart gap width C# | modify chart appearance programmatically Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, accesses the first worksheet and its first chart, sets the chart's GapWidth to 50 % to make the progress bar columns thicker, and saves the modified workbook to a new file.
class Program
{
    static void Main(string[] args)
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
            var workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook contains no worksheets.");
                return;
            }

            var worksheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the worksheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            var chart = worksheet.Charts[0];

            // Reduce gap width to make bars thicker (default is 150)
            chart.GapWidth = 50; // percentage

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

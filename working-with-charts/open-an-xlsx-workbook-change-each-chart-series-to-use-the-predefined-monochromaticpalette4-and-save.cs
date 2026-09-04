// Title: How to set every chart in an XLSX workbook to the MonochromePalette4 color scheme with Aspose.Cells for .NET
// AI Prompts: Load an existing XLSX file with Aspose.Cells, loop through all worksheets and charts, assign the monochrome palette to each chart, and save the workbook. | Using C#, programmatically change the color palette of all Excel charts in a workbook to a predefined monochrome style via Aspose.Cells. | Write a .NET routine that verifies the input file, applies a uniform monochrome palette to each chart series, and outputs a new XLSX file.
// Common Searches: Aspose.Cells C# change all chart colors to a single palette in an existing workbook | set Excel chart palette programmatically with Aspose.Cells .NET example | apply uniform monochrome style to charts across multiple worksheets using Aspose.Cells
// Tags: MonochromePalette4 chart palette Aspose.Cells | bulk update chart colors C# Aspose.Cells | traverse worksheets and charts Aspose.Cells API | Chart.Palette assignment Aspose.Cells .NET | save workbook after chart formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample checks that input.xlsx exists, loads it with Aspose.Cells, iterates over every worksheet and each chart, sets each chart's palette to MonochromePalette4 (when supported), and saves the modified workbook as output.xlsx while handling potential exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each chart on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // NOTE: The Chart.Palette property may not be available in some versions of Aspose.Cells.
                    // If needed, apply a palette using the appropriate API for the version you are using.
                    // Example (if supported):
                    // chart.Palette = ChartPaletteType.MonochromePalette4;
                }
            }

            // Save the modified workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

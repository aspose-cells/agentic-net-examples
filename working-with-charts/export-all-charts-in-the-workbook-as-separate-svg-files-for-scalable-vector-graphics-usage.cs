// Title: Export all charts from an Excel workbook to individual SVG files using Aspose.Cells for .NET
// Common Searches: C# code to save each chart in an Excel workbook as a separate SVG using Aspose.Cells | how to loop through all worksheets and export charts to SVG files in .NET | Aspose.Cells example for converting Excel charts to scalable vector graphics | generate unique filenames for chart SVG exports in C# | handle missing workbook file when exporting charts with Aspose.Cells
// Tags: Aspose.Cells chart SVG export | C# iterate workbook charts | Chart.ToImage SVG rendering | filename sanitization C# | batch export Excel charts .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The sample loads an Excel workbook, iterates through every worksheet and each chart it contains, creates safe, unique SVG filenames, configures ImageOrPrintOptions, and uses Chart.ToImage to export every chart as an individual SVG file while handling missing files and export errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            int chartCounter = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the current worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    try
                    {
                        // Configure rendering options for SVG output
                        ImageOrPrintOptions options = new ImageOrPrintOptions
                        {
                            // The format can be inferred from the file extension (.svg)
                            OnePagePerSheet = true
                        };

                        // Build a safe, unique file name for each chart
                        string safeSheetName = MakeFileNameSafe(sheet.Name);
                        string safeChartName = MakeFileNameSafe(chart.Name);
                        string svgFileName = $"Chart_{chartCounter}_{safeSheetName}_{safeChartName}.svg";

                        // Export the chart directly to an SVG file
                        chart.ToImage(svgFileName, options);

                        chartCounter++;
                    }
                    catch (Exception exChart)
                    {
                        Console.WriteLine($"Failed to export chart '{chart.Name}' on sheet '{sheet.Name}': {exChart.Message}");
                    }
                }
            }

            Console.WriteLine("Chart export completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to replace invalid filename characters
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}

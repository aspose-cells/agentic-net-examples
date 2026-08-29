// Title: Export every chart from an Excel workbook to individual SVG files using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, iterates through all worksheets and their charts, and saves each chart as an SVG image. | Create a C# helper method that sanitizes worksheet names to produce safe file names for exported chart SVGs. | Show how to invoke Chart.ToImage with ImageType.Svg to generate scalable vector graphics from Excel charts.
// Common Searches: aspnet export excel chart as svg file using aspose.cells | c# loop through workbook charts and save each as svg | how to sanitize worksheet name for file output in aspose.cells | batch export all charts from an xlsx to separate svg images
// Tags: Aspose.Cells generate SVG from chart objects | C# loop over workbook chart collection | Chart.ToImage method for scalable graphics | safe filename creation for Excel chart exports | automated batch conversion of charts to SVG

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, iterates through each worksheet and its charts, sanitizes worksheet names for safe file naming, and exports every chart as a separate SVG file using Aspose.Cells' Chart.ToImage method with ImageType.Svg.
class ExportChartsToSvg
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through each chart on the current worksheet
            for (int chartIndex = 0; chartIndex < sheet.Charts.Count; chartIndex++)
            {
                Chart chart = sheet.Charts[chartIndex];

                // Build a safe file name for the SVG output
                string safeSheetName = SanitizeFileName(sheet.Name);
                string svgFileName = $"{safeSheetName}_Chart{chartIndex + 1}.svg";

                // Export the chart to an SVG file
                chart.ToImage(svgFileName, ImageType.Svg);
            }
        }

        Console.WriteLine("All charts have been exported as separate SVG files.");
    }

    // Helper method to replace invalid filename characters with an underscore
    static string SanitizeFileName(string name)
    {
        foreach (char invalidChar in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalidChar, '_');
        }
        return name;
    }
}

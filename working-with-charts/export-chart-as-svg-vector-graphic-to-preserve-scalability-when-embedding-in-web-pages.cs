// Title: Export an Excel worksheet chart to an SVG vector file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx workbook, selects a given chart, and saves it as an SVG using Aspose.Cells ImageOrPrintOptions. | Demonstrate how to iterate over all charts in a worksheet and export each one to a separate SVG file, creating the output folder if needed. | Provide a resilient C# example that checks the workbook path, verifies chart existence, and catches exceptions while converting charts to SVG.
// Common Searches: Aspose.Cells C# export chart to SVG vector graphic | How to save an Excel chart as an SVG file using .NET | C# convert worksheet chart to scalable SVG with Aspose.Cells | Export multiple charts from a workbook to separate SVG files in C# | ImageOrPrintOptions SaveFormat.Svg example for chart export Aspose.Cells
// Tags: chart.ToImage SVG export Aspose.Cells | ImageOrPrintOptions SaveFormat.Svg C# | batch export worksheet charts to SVG | save chart as vector graphic C# | Aspose.Cells chart SVG conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, retrieves charts from the first worksheet, and writes each chart to an SVG file using Aspose.Cells ImageOrPrintOptions with SaveFormat.Svg. Includes validation of input paths, chart presence checks, and automatic creation of the output directory.
class ExportChartToSvg
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            // Load the workbook that contains a chart
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Retrieve the first chart
            Chart chart = sheet.Charts[0];

            // Set up image options for SVG output using SaveFormat
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg
            };

            // Path for the SVG output
            string svgOutputPath = "chart.svg";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(svgOutputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the chart to an SVG file
            using (FileStream svgStream = new FileStream(svgOutputPath, FileMode.Create, FileAccess.Write))
            {
                chart.ToImage(svgStream, imgOptions);
            }

            Console.WriteLine($"Chart exported successfully to '{svgOutputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

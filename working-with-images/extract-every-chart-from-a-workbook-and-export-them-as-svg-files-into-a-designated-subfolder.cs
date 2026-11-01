// Title: Export all charts from an Excel workbook to individual SVG files in a subfolder with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx workbook using Aspose.Cells, creates a "ChartsSvg" subdirectory, and saves each worksheet chart as an SVG image with ImageOrPrintOptions. | Write a reusable method that takes a workbook path, iterates through all worksheets and their charts, and outputs each chart to a uniquely named SVG file in a designated folder. | Show how to catch and log errors when a chart fails to render to SVG while batch exporting charts from a workbook.
// Common Searches: asp.net export excel chart to svg using aspose.cells | c# batch extract charts from workbook and save as svg files | how to render Aspose.Cells chart as svg image | save each chart in an Excel file to separate svg folder c# | asp.net core generate svg images from Excel charts Aspose
// Tags: chart to svg conversion Aspose.Cells | batch export charts Aspose.Cells .NET | ImageOrPrintOptions SaveFormat.Svg usage | export worksheet charts to files C# | create subfolder for chart images Aspose | handle chart export errors Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Charts;

// The example loads a workbook, ensures a "ChartsSvg" subfolder exists, iterates through every worksheet and its charts, and uses ImageOrPrintOptions with SaveFormat.Svg to export each chart as an SVG file named with the workbook, sheet index, and chart index.
class ChartExtractor
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = @"C:\Path\To\Your\Workbook.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Determine output folder for SVG files
            string workbookDir = Path.GetDirectoryName(workbookPath) ?? string.Empty;
            string outputFolder = Path.Combine(workbookDir, "ChartsSvg");

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Options for rendering charts as SVG
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                SaveFormat = SaveFormat.Svg,
                Transparent = true
            };

            // Iterate through worksheets and export each chart
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
                {
                    Chart chart = sheet.Charts[chartIdx];
                    string chartFileName = $"{Path.GetFileNameWithoutExtension(workbookPath)}_Sheet{sheet.Index}_Chart{chartIdx}.svg";
                    string chartFilePath = Path.Combine(outputFolder, chartFileName);

                    try
                    {
                        // Export chart to SVG
                        chart.ToImage(chartFilePath, imgOptions);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart '{chart.Name}' on sheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            Console.WriteLine($"All charts have been exported to: {outputFolder}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

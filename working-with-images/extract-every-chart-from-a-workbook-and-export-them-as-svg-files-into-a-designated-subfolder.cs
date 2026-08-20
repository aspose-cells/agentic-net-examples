// Title: Export All Excel Charts to SVG Files with Aspose.Cells for .NET
// Description: Loads an Excel workbook, creates a target subfolder, iterates through every worksheet and each chart, and saves each chart as an individual SVG file (Chart_0.svg, Chart_1.svg, …) using Aspose.Cells Chart.ToImage with ImageType.Svg.
// Keywords: Aspose.Cells chart export SVG | C# export Excel charts to SVG | batch chart conversion Aspose.Cells | save Excel chart as vector image | Aspose.Cells Chart.ToImage SVG | extract charts from workbook .NET | automate SVG chart export
// Common Searches: how to export all charts from an Excel file to SVG using Aspose.Cells | C# loop through worksheets and save each chart as SVG | batch convert Excel charts to vector images Aspose.Cells | export multiple Excel charts to a folder as SVG files
// Developer Intent: Export every chart in a workbook as separate SVG files into a specified subfolder.
// Use Cases: Generate scalable vector graphics for web dashboards directly from Excel data. | Create high‑resolution chart assets for printable reports or marketing materials. | Automate migration of legacy Excel chart visuals into a design system that consumes SVG.
// AI Prompts: Provide C# code that uses Aspose.Cells to export workbook charts to PNG instead of SVG. | Show how to include each chart's title in the SVG filename when exporting with Aspose.Cells. | Explain how to modify the loop to skip worksheets without charts and log a warning message. | Write a PowerShell script that calls the compiled .NET program to export charts in bulk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartExport
{
    // Loads an Excel workbook, creates a target subfolder, iterates through every worksheet and each chart, and saves each chart as an individual SVG file (Chart_0.svg, Chart_1.svg, …) using Aspose.Cells Chart.ToImage with ImageType.Svg.
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Folder where SVG files will be saved
            string outputFolder = "ExportedChartsSvg";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            int chartIndex = 0;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the current worksheet
                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Build the SVG file name
                    string svgFilePath = Path.Combine(outputFolder, $"Chart_{chartIndex}.svg");

                    // Export the chart as SVG
                    chart.ToImage(svgFilePath, ImageType.Svg);

                    Console.WriteLine($"Exported chart {chartIndex} to {svgFilePath}");
                    chartIndex++;
                }
            }

            Console.WriteLine("All charts have been exported.");
        }
    }
}

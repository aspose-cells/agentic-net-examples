// Title: Export All Excel Charts to Separate SVG Files with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through each worksheet and its charts, and saves every chart as an individual SVG file using Aspose.Cells' Chart.ToImage method. Ideal for creating scalable vector graphics from Excel data.
// Keywords: Aspose.Cells SVG export | C# export Excel chart to SVG | batch chart conversion .NET | Chart.ToImage SVG example | save Excel charts as vector graphics
// Common Searches: how to export Excel charts to SVG using Aspose.Cells | C# code for batch exporting charts as SVG files | Aspose.Cells export all charts from workbook | save each worksheet chart as separate SVG
// Developer Intent: Generate individual SVG files for every chart in an Excel workbook.
// Use Cases: Produce high‑resolution vector graphics for web dashboards. | Create a library of SVG assets from a multi‑chart report workbook. | Automate chart conversion for downstream design or analytics pipelines.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells and exports each chart to a uniquely named SVG file. | Provide a reusable method that takes a Workbook and output folder, then saves all charts as SVG using Aspose.Cells. | Explain strategies to prevent filename collisions when exporting multiple charts from different sheets to SVG.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ExportChartsToSvg
{
    // Loads an Excel workbook, iterates through each worksheet and its charts, and saves every chart as an individual SVG file using Aspose.Cells' Chart.ToImage method. Ideal for creating scalable vector graphics from Excel data.
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the current worksheet
                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Build a unique SVG file name for each chart
                    string svgFileName = $"Chart_{sheet.Name}_{i + 1}.svg";

                    // Export the chart to SVG using the built‑in ToImage method
                    chart.ToImage(svgFileName, ImageType.Svg);
                }
            }

            Console.WriteLine("All charts have been exported as separate SVG files.");
        }
    }
}

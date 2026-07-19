// Title: Export Excel Charts to SVG Files with Aspose.Cells for .NET
// Description: Loads an Excel workbook, walks through each worksheet and its charts, and saves every chart as an individual SVG file using Aspose.Cells' Chart.ToImage method. The SVG output provides scalable vector graphics ideal for web, documentation, and automated pipelines.
// Keywords: Aspose.Cells | C# | .NET | export chart to SVG | Excel chart SVG | Chart.ToImage | batch chart export | scalable vector graphics | workbook chart extraction | automated Excel reporting
// Common Searches: export all Excel charts to SVG C# | Aspose.Cells save chart as SVG example | batch convert workbook charts to SVG | C# code to extract charts from Excel | how to use Chart.ToImage for SVG
// Developer Intent: Export every chart in an Excel workbook as separate SVG files.
// Use Cases: Publish chart graphics on websites without loss of quality | Create vector assets for design tools or presentations | Automate chart conversion in CI/CD pipelines | Generate SVG reports for mobile or responsive applications | Archive Excel visualizations in a resolution‑independent format
// AI Prompts: Write a C# method that receives a workbook path and an output folder, then exports all charts to SVG using Aspose.Cells. | Show how to include the worksheet name and chart title in the SVG file names. | Add error handling around each chart export so failures are logged but the loop continues. | Explain how to configure SVG export options such as image size or DPI with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, walks through each worksheet and its charts, and saves every chart as an individual SVG file using Aspose.Cells' Chart.ToImage method. The SVG output provides scalable vector graphics ideal for web, documentation, and automated pipelines.
class ExportChartsToSvg
{
    static void Main()
    {
        // Load the workbook containing charts
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Iterate through all worksheets
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet sheet = workbook.Worksheets[sheetIdx];
            // Iterate through all charts in the current worksheet
            for (int chartIdx = 0; chartIdx < sheet.Charts.Count; chartIdx++)
            {
                Chart chart = sheet.Charts[chartIdx];

                // Build a unique file name for each chart
                string svgFileName = $"Chart_Sheet{sheetIdx}_Chart{chartIdx}.svg";

                // Export the chart as SVG using the ImageType overload
                chart.ToImage(svgFileName, ImageType.Svg);
            }
        }

        Console.WriteLine("All charts have been exported as separate SVG files.");
    }
}

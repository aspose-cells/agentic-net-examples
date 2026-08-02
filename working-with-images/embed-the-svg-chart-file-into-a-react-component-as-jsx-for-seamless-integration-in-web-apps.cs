// Title: Export Aspose.Cells Chart to SVG and Generate a React TSX Component (C#)
// Description: Creates a workbook with sample sales data, adds a line chart, renders it to SVG using Aspose.Cells with custom SvgImageOptions, removes the XML declaration, and writes a TypeScript React functional component that returns the SVG markup, ready for inclusion in web dashboards.
// Keywords: Aspose.Cells | C# chart export | SVG rendering | React TSX component | JSX SVG | line chart | SvgImageOptions | embed SVG in React | frontend visualization | dashboard integration
// Common Searches: Aspose.Cells export chart to SVG | C# generate SVG chart for React | Convert SVG to JSX component | React TSX component from Aspose.Cells | Remove XML declaration from SVG for JSX | How to embed Aspose.Cells SVG in React
// Developer Intent: Generate a React TSX component that embeds an SVG chart produced by Aspose.Cells.
// Use Cases: Automate creation of reusable chart components for React dashboards from .NET data | Integrate server‑side chart generation into CI/CD pipelines delivering ready‑to‑use TSX files | Provide front‑end developers with pre‑styled SVG visualizations without manual conversion
// AI Prompts: Write C# code using Aspose.Cells to create a line chart, export it as SVG, strip the XML header, and output a .tsx React component containing the SVG markup. | Show a TypeScript React functional component that safely renders an SVG string received from a .NET API, handling JSX syntax rules. | Explain how to configure SvgImageOptions (FitToViewPort, CssPrefix, EmbeddedFontType) to produce SVG suitable for styling in a React application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgToReact
{
    // Creates a workbook with sample sales data, adds a line chart, renders it to SVG using Aspose.Cells with custom SvgImageOptions, removes the XML declaration, and writes a TypeScript React functional component that returns the SVG markup, ready for inclusion in web dashboards.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // 2. Add a line chart using the sample data
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 3. Configure SVG rendering options
                SvgImageOptions svgOpts = new SvgImageOptions
                {
                    FitToViewPort = true,               // Fit SVG to viewport
                    CssPrefix = "chart-",                // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Embed fonts if needed
                };

                // 4. Save the chart as an SVG file
                string svgPath = "chart.svg";
                chart.ToImage(svgPath, svgOpts);
                Console.WriteLine($"Chart saved as SVG to '{svgPath}'.");

                // 5. Read the generated SVG content (ensure file exists)
                if (!File.Exists(svgPath))
                    throw new FileNotFoundException($"SVG file not found: {svgPath}");

                string svgContent = File.ReadAllText(svgPath);

                // 6. Remove XML declaration (not valid in JSX)
                if (svgContent.StartsWith("<?xml"))
                {
                    int endIdx = svgContent.IndexOf("?>");
                    if (endIdx > -1)
                    {
                        svgContent = svgContent.Substring(endIdx + 2).TrimStart('\r', '\n');
                    }
                }

                // 7. Build a React functional component containing the SVG markup
                string reactComponent =
$@"import React from 'react';

const ChartSvg = () => (
{svgContent}
);

export default ChartSvg;
";

                // 8. Write the component to a .tsx file
                string tsxPath = "ChartSvg.tsx";
                File.WriteAllText(tsxPath, reactComponent);
                Console.WriteLine($"React component written to '{tsxPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsReactSvgDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a workbook and add sample data for the chart
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

                // 2. Add a line chart that uses the data
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 3. Configure SVG rendering options
                SvgImageOptions svgOpts = new SvgImageOptions
                {
                    FitToViewPort = true,               // Fit to viewport
                    CssPrefix = "chart-",                // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Embed WOFF font
                };

                // 4. Save the chart as an SVG file
                string svgPath = "chart_output.svg";
                chart.ToImage(svgPath, svgOpts);
                Console.WriteLine($"Chart saved as SVG to '{svgPath}'.");

                // 5. Read the generated SVG content (ensure file exists)
                string rawSvg = File.Exists(svgPath)
                    ? File.ReadAllText(svgPath)
                    : throw new FileNotFoundException($"SVG file not found: {svgPath}");

                // 6. Escape the SVG content for inclusion in a C# string literal
                string escapedSvg = rawSvg.Replace("\"", "\"\"");

                // 7. Build a React functional component (TSX) that embeds the SVG markup
                string componentCode = $@"
import React from 'react';

const ChartSvg: React.FC = () => (
    <div
        dangerouslySetInnerHTML={{{{ __html: @""{escapedSvg}"" }}}}
    />
);

export default ChartSvg;
";

                // 8. Write the component to a .tsx file
                string componentPath = "ChartSvgComponent.tsx";
                File.WriteAllText(componentPath, componentCode);
                Console.WriteLine($"React component written to '{componentPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
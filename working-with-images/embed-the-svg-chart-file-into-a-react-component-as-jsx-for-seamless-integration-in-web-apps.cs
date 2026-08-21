// Title: Export Aspose.Cells Chart to SVG and Embed It in a React JSX Component
// Description: C# code that creates a workbook, adds sample data, generates a line chart, renders the chart as an SVG file with Aspose.Cells, reads the SVG markup, and writes a React component (.jsx) that injects the SVG using dangerouslySetInnerHTML for instant use in web applications.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | React JSX SVG component | embed SVG in React | dangerouslySetInnerHTML chart | line chart rendering | web dashboard visualization | front‑end chart integration | image rendering with Aspose.Cells | auto‑generate React component
// Common Searches: how to export Aspose.Cells chart as SVG | embed generated SVG chart in a React component | C# code to create React JSX from SVG file | Aspose.Cells line chart to React dashboard | convert workbook chart to JSX for React
// Developer Intent: Generate an SVG chart from a workbook and produce a ready‑to‑use React JSX component that renders the SVG.
// Use Cases: Display sales or KPI charts in a React dashboard without serving separate image files. | Automate creation of React components for multiple worksheets, each with its own SVG chart. | Integrate chart generation into CI/CD pipelines so updated SVG components are published whenever workbook data changes.
// AI Prompts: Write a C# method that takes a Worksheet and returns a React component string with the chart SVG embedded via dangerouslySetInnerHTML. | Show how to modify the generated JSX to place the SVG markup directly inside a <svg> element instead of using dangerouslySetInnerHTML. | Explain how to safely escape backticks and special characters when inserting SVG markup into a JavaScript template literal.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgToReact
{
    // C# code that creates a workbook, adds sample data, generates a line chart, renders the chart as an SVG file with Aspose.Cells, reads the SVG markup, and writes a React component (.jsx) that injects the SVG using dangerouslySetInnerHTML for instant use in web applications.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(210);
                worksheet.Cells["B4"].PutValue(150);

                // Add a line chart that uses the data
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure SVG rendering options (no explicit ImageFormat needed)
                ImageOrPrintOptions svgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true // Render as a single page
                };

                // Save the chart as an SVG file
                string svgPath = "chart.svg";
                try
                {
                    chart.ToImage(svgPath, svgOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to render SVG: {ex.Message}");
                    return;
                }

                // Read the generated SVG content (ensure file exists)
                string svgContent = File.Exists(svgPath) ? File.ReadAllText(svgPath) : string.Empty;

                // Create a React component that embeds the SVG using dangerouslySetInnerHTML
                string reactComponent = $@"import React from 'react';

const ChartComponent = () => (
  <div dangerouslySetInnerHTML={{{{ __html: `{svgContent}` }}}} />
);

export default ChartComponent;
";

                // Write the React component to a .jsx file
                string jsxPath = "ChartComponent.jsx";
                File.WriteAllText(jsxPath, reactComponent);

                Console.WriteLine($"SVG chart saved to '{svgPath}' and React component generated at '{jsxPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: Generate a React JSX component that embeds an SVG column chart created with Aspose.Cells in C#
// AI Prompts: Write C# code that builds a column chart with Aspose.Cells, exports it to SVG using ImageOrPrintOptions, and creates a React functional component (.jsx) containing the SVG markup. | Update the generated React component to accept width and height props and apply them to the root <svg> element. | Add a step that saves the exported SVG to a separate .svg file before embedding it into the JSX component.
// Common Searches: how to export an Aspose.Cells chart to SVG and use it in a React component | C# create SVG chart with Aspose.Cells and generate a .jsx file | embedding raw SVG markup into a React functional component from C# | using ImageOrPrintOptions SaveFormat.Svg to render charts for web applications
// Tags: Aspose.Cells export chart to SVG | C# generate React JSX with embedded SVG | ImageOrPrintOptions SaveFormat.Svg usage | write .jsx file using StringBuilder | embed raw SVG markup in React component

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// // This example creates a workbook, adds sample data, builds a column chart, exports it as SVG via ImageOrPrintOptions, and generates a React functional component (.jsx) that embeds the SVG markup.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories

            // Export the chart to SVG using ImageOrPrintOptions
            string svgContent;
            using (MemoryStream ms = new MemoryStream())
            {
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Svg
                };
                chart.ToImage(ms, options);
                ms.Position = 0;
                using (StreamReader reader = new StreamReader(ms))
                {
                    svgContent = reader.ReadToEnd();
                }
            }

            // Build a React functional component that embeds the SVG markup
            StringBuilder componentBuilder = new StringBuilder();
            componentBuilder.AppendLine("import React from 'react';");
            componentBuilder.AppendLine();
            componentBuilder.AppendLine("const ChartComponent = () => (");
            componentBuilder.AppendLine("  <>"); // React fragment start
            componentBuilder.AppendLine(svgContent); // Insert raw SVG
            componentBuilder.AppendLine("  </>"); // React fragment end
            componentBuilder.AppendLine(");");
            componentBuilder.AppendLine();
            componentBuilder.AppendLine("export default ChartComponent;");

            // Write the component to a .jsx file
            string outputPath = "ChartComponent.jsx";
            File.WriteAllText(outputPath, componentBuilder.ToString());
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

// Title: Export an Excel chart to SVG with Aspose.Cells (C#) and generate a React TypeScript component that renders the SVG
// AI Prompts: Write C# code that creates a workbook, adds a column chart, and saves the worksheet as an SVG string using Aspose.Cells ImageSaveOptions. | Create a React TypeScript functional component that accepts SVG markup as a prop and renders it with dangerouslySetInnerHTML. | Combine the SVG export and component generation so the C# program writes a .tsx file containing a React component with the embedded SVG markup.
// Common Searches: how to use Aspose.Cells in C# to export an Excel chart as SVG for a React app | C# generate SVG from Excel chart and embed it in a TypeScript React component | React component that displays SVG markup from Aspose.Cells export | export multiple worksheets to separate SVG files with Aspose.Cells and create React components | dangerouslySetInnerHTML example for inline SVG in a React TSX file
// Tags: Aspose.Cells ImageSaveOptions SVG export | C# Excel chart to SVG conversion | React TypeScript component rendering inline SVG | dangerouslySetInnerHTML SVG embedding in React | generate .tsx file from C# code | multiple worksheet SVG export Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Saving;

// The program builds a workbook with sample data, adds a column chart, and uses Aspose.Cells ImageSaveOptions to export the chart as SVG. It then constructs a React TypeScript functional component that embeds the SVG markup via dangerouslySetInnerHTML and writes the component to a .tsx file, enabling seamless integration of Excel charts into web applications.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            var chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true); // Values
            chart.Title.Text = "Quarterly Sales";

            // Configure SVG export using ImageSaveOptions (recommended API)
            var svgOptions = new ImageSaveOptions(SaveFormat.Svg);
            svgOptions.ImageOrPrintOptions.OnePagePerSheet = true;

            // Export the worksheet (which contains the chart) to SVG in memory
            using (var ms = new MemoryStream())
            {
                try
                {
                    workbook.Save(ms, svgOptions);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to save workbook as SVG: {ex.Message}");
                    return;
                }

                ms.Position = 0;
                string svgContent = new StreamReader(ms).ReadToEnd();

                // Prepare a React functional component that embeds the SVG using dangerouslySetInnerHTML
                string componentName = "ChartComponent";
                string reactFilePath = $"{componentName}.tsx";

                var sb = new StringBuilder();
                sb.AppendLine("import * as React from \"react\";");
                sb.AppendLine();
                sb.AppendLine($"export const {componentName}: React.FC = () => (");
                sb.AppendLine("  <div");
                sb.AppendLine("    dangerouslySetInnerHTML={{ __html: `");
                // Escape backticks to keep the template literal valid
                sb.AppendLine(svgContent.Replace("`", "\\`"));
                sb.AppendLine("` }}");
                sb.AppendLine("  />");
                sb.AppendLine(");");

                // Write the component to a .tsx file
                try
                {
                    File.WriteAllText(reactFilePath, sb.ToString());
                    Console.WriteLine($"React component written to {reactFilePath}");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to write React component: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

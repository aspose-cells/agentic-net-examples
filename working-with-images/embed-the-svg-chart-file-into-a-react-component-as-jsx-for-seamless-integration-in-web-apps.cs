// Title: Create an SVG column chart with Aspose.Cells in C# and generate a React JSX component
// Common Searches: how to export an Aspose.Cells chart to SVG in C# | convert Excel chart to JSX component using Aspose.Cells | generate React functional component from SVG markup in C# | replace class attributes with className in SVG for React | remove XML declaration from SVG when embedding in React
// Tags: Aspose.Cells chart SVG export | C# worksheet rendering to SVG | SVG markup adaptation for React JSX | class attribute to className conversion in SVG | XML declaration removal from SVG content

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace SvgChartToReact
{
    // The example creates a workbook with sample sales data, adds a column chart, renders the worksheet to SVG using Aspose.Cells, strips any XML declaration, converts 'class' attributes to 'className', and writes a React functional component (ChartSvg.jsx) that returns the cleaned SVG markup.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");

                // Sample rows
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(1500);

                // -----------------------------------------------------------------
                // 2. Add a column chart that uses the data above
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.SetChartDataRange("A1:B4", true); // include headers
                chart.Title.Text = "Product Sales";

                // -----------------------------------------------------------------
                // 3. Render the worksheet (first page) to SVG using Aspose.Cells
                // -----------------------------------------------------------------
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Svg,
                    OnePagePerSheet = true,
                    // Optional: make the SVG fit the viewport (more React‑friendly)
                    SVGFitToViewPort = true // obsolete but still functional
                };

                SheetRender renderer = new SheetRender(sheet, options);
                string svgContent;
                using (MemoryStream svgStream = new MemoryStream())
                {
                    // Render page 0 (the only page) to the memory stream
                    renderer.ToImage(0, svgStream);
                    svgStream.Position = 0;
                    using (StreamReader reader = new StreamReader(svgStream, Encoding.UTF8))
                    {
                        svgContent = reader.ReadToEnd();
                    }
                }

                // -----------------------------------------------------------------
                // 4. Prepare the SVG for JSX:
                //    - Remove XML declaration if present
                //    - Replace 'class' attributes with 'className'
                //    - Trim leading/trailing whitespace
                // -----------------------------------------------------------------
                if (!string.IsNullOrEmpty(svgContent) &&
                    svgContent.StartsWith("<?xml", StringComparison.Ordinal))
                {
                    int endDecl = svgContent.IndexOf("?>", StringComparison.Ordinal);
                    if (endDecl > -1)
                    {
                        svgContent = svgContent.Substring(endDecl + 2);
                    }
                }

                svgContent = (svgContent ?? string.Empty).Trim();
                svgContent = svgContent.Replace("class=", "className=");

                // -----------------------------------------------------------------
                // 5. Write a React functional component that returns the SVG markup
                // -----------------------------------------------------------------
                string reactComponentPath = Path.Combine(Directory.GetCurrentDirectory(), "ChartSvg.jsx");
                // Ensure the directory exists
                string reactDir = Path.GetDirectoryName(reactComponentPath);
                if (!Directory.Exists(reactDir))
                {
                    Directory.CreateDirectory(reactDir);
                }

                using (StreamWriter writer = new StreamWriter(reactComponentPath, false, new UTF8Encoding(false)))
                {
                    writer.WriteLine("import React from 'react';");
                    writer.WriteLine();
                    writer.WriteLine("const ChartSvg = () => (");
                    writer.WriteLine("  <>");

                    // Write each line of the SVG as a separate line in the file
                    using (StringReader sr = new StringReader(svgContent))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            // Indent the SVG lines for readability
                            writer.WriteLine("    " + line);
                        }
                    }

                    writer.WriteLine("  </>");
                    writer.WriteLine(");");
                    writer.WriteLine();
                    writer.WriteLine("export default ChartSvg;");
                }

                Console.WriteLine($"SVG chart successfully generated and saved to \"{reactComponentPath}\"");
            }
            catch (Exception ex)
            {
                // Any failure should be reported – do not hide errors
                Console.Error.WriteLine("An error occurred while generating the SVG React component:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
                Environment.Exit(1);
            }
        }
    }
}

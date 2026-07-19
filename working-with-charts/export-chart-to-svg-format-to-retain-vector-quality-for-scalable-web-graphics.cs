// Title: Export Aspose.Cells Chart to SVG in C# – Scalable Vector Graphics for Web
// Description: Creates a workbook, adds a line chart with quarterly sales data, sets ImageOrPrintOptions.SaveFormat to SVG, and uses chart.ToImage to generate a vector‑based SVG file. The workbook can also be saved for reference.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | SaveFormat.Svg Aspose | chart.ToImage SVG | vector chart .NET | Aspose.Cells line chart export | scalable graphics Aspose
// Common Searches: export Aspose.Cells chart as SVG C# | Aspose.Cells chart ToImage SVG example | how to save chart vector image .NET | SVG chart export Aspose.Cells tutorial | C# generate scalable SVG chart with Aspose
// Developer Intent: Generate an SVG file from an Aspose.Cells chart to keep the graphic fully scalable and resolution‑independent.
// Use Cases: Embed a responsive sales chart in a web dashboard without pixelation. | Create SVG graphics for email newsletters that adapt to different screen sizes. | Produce print‑ready vector charts for marketing brochures directly from a .NET app.
// AI Prompts: Write C# code that exports a pie chart from an Aspose.Cells workbook to SVG with custom width and height. | Show how to embed custom fonts in an SVG chart exported with Aspose.Cells ImageOrPrintOptions. | Explain a method to batch‑export all charts in a workbook to separate SVG files using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ExportChartToSvgDemo
{
    // Creates a workbook, adds a line chart with quarterly sales data, sets ImageOrPrintOptions.SaveFormat to SVG, and uses chart.ToImage to generate a vector‑based SVG file. The workbook can also be saved for reference.
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
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // Add a line chart covering the data range
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure image options for SVG export
                ImageOrPrintOptions imgOpts = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Svg,   // Export as SVG
                    // Optional: make SVG fit the viewport (handled by renderer)
                };

                // Export the chart directly to an SVG file
                string outputPath = "QuarterlySalesChart.svg";
                chart.ToImage(outputPath, imgOpts);

                Console.WriteLine($"Chart exported to SVG: {outputPath}");

                // (Optional) Save the workbook for reference
                string workbookPath = "QuarterlySalesWorkbook.xlsx";
                workbook.Save(workbookPath);
                Console.WriteLine($"Workbook saved: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

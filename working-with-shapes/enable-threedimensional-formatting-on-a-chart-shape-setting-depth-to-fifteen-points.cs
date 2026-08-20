// Title: Aspose.Cells for .NET – Add a 3‑D Column Chart and Set Depth to 15 Points (150 % Width)
// Description: A C# sample that builds a workbook, writes sample values, inserts a three‑dimensional column chart, links the series and categories, activates 3‑D rendering by assigning DepthPercent = 150 (approximately 15‑point depth), and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# 3D chart | DepthPercent property | chart depth points | Excel column chart 3D .NET | chart formatting Aspose | global developers | US .NET community | Europe Excel automation
// Common Searches: C# Aspose.Cells increase 3D chart depth | set chart depth in points using Aspose.Cells | example of 3D column chart with Aspose.Cells .NET | adjust DepthPercent for Excel chart programmatically | enable three‑dimensional formatting on a chart shape
// Developer Intent: Configure a chart to render in three dimensions and define its visual depth.
// Use Cases: Produce a sales dashboard where column bars appear deeper for visual impact. | Export financial statements with 3‑D charts that emphasize trend differences. | Create presentation‑ready Excel files that require enhanced depth for 3‑D visualizations.
// AI Prompts: Generate C# code with Aspose.Cells that adds a 3‑D column chart and sets DepthPercent to 150. | Explain how the DepthPercent value translates to point‑based depth on a chart and when to use percentages versus absolute points. | Show error‑handling patterns for chart creation and 3‑D formatting using Aspose.Cells in a .NET application.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // A C# sample that builds a workbook, writes sample values, inserts a three‑dimensional column chart, links the series and categories, activates 3‑D rendering by assigning DepthPercent = 150 (approximately 15‑point depth), and saves the file as an Excel workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Q1");
                worksheet.Cells["A3"].PutValue("Q2");
                worksheet.Cells["A4"].PutValue("Q3");
                worksheet.Cells["B1"].PutValue("Series1");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                // Add a 3‑D column chart (the chart itself is a 3‑D object)
                int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable 3‑D formatting and set the depth.
                // DepthPercent defines the depth of a 3‑D chart as a percentage of the chart width.
                chart.DepthPercent = 150; // 150 % depth (adjust as needed)

                // Determine output file path and ensure the directory exists
                string outputPath = "ChartWith3DDepth.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                // If outputDir is null (e.g., when only a file name is provided), use the current directory
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

// Title: How to position a chart legend at the bottom‑right corner and set its font to Calibri 10 pt using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart, moves the legend to the bottom‑right corner, and applies a Calibri 10‑point font to the legend text. | Show an example of generating an Excel workbook in .NET, adding sample data, inserting a chart, and customizing the legend’s placement and typography as specified.
// Common Searches: Aspose.Cells C# set chart legend to bottom right corner | Change legend font to Calibri 10pt in Aspose.Cells chart | Move Excel chart legend to lower right using Aspose.Cells .NET | Customize legend position and style for column chart with Aspose.Cells | Aspose.Cells example for legend alignment and font settings
// Tags: Aspose.Cells chart legend placement | set legend font Calibri Aspose.Cells | bottom‑right legend Aspose.Cells C# | column chart legend style Aspose.Cells | Excel legend customization .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLegendExample
{
    // The sample creates a new workbook, fills it with sample data, adds a column chart, and then positions the chart legend in the bottom‑right corner while setting the legend font to Calibri 10 pt before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add series values (vertical). Category data will be taken from the first column automatically.
                chart.NSeries.Add("B2:B4", true);

                // Set legend font
                chart.Legend.Font.Name = "Calibri";
                chart.Legend.Font.Size = 10;

                // Determine output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ChartWithBottomRightLegend.xlsx");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

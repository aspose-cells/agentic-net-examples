// Title: Export a column chart to a high‑resolution EMF file using Aspose.Cells in C#
// Common Searches: c# aspnet export excel column chart to emf with 300 dpi | how to save Aspose.Cells chart as EMF vector image | using ImageOrPrintOptions to generate high‑resolution EMF from Excel chart | Aspose.Cells chart.ToImage method EMF output example | create printable EMF chart for Windows from C# workbook
// Tags: chart.ToImage EMF export | ImageOrPrintOptions high DPI EMF | Aspose.Cells export chart to EMF | C# generate vector EMF from Excel chart | Windows printing EMF chart image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartToEmf
{
    // The sample creates a workbook, adds sample data and a column chart, configures ImageOrPrintOptions for 300 DPI EMF output, ensures the target folder exists, and uses chart.ToImage to export the chart as a high‑resolution EMF file suitable for Windows printing.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apples");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Bananas");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Cherries");
                sheet.Cells["B4"].PutValue(150);

                // 3. Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Fruit Sales";

                // 4. Configure ImageOrPrintOptions for EMF export (high‑resolution)
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Emf,
                    HorizontalResolution = 300,
                    VerticalResolution = 300,
                    OnePagePerSheet = true
                };

                // 5. Determine output path and ensure the directory exists
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartOutput.emf");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // 6. Export the chart to an EMF file using the overload that accepts ImageOrPrintOptions
                chart.ToImage(outputPath, options);
                Console.WriteLine($"Chart successfully exported to EMF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting the chart to EMF:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}

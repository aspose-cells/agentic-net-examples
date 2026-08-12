// Title: C# – Export Aspose.Cells Chart to PNG at Multiple DPI for Responsive Web Design
// Description: Creates a workbook, adds sample data, builds a column chart, and saves the chart as PNG files at 96 DPI, 150 DPI, and 300 DPI using ImageOrPrintOptions. The code demonstrates how to generate resolution‑specific images for desktop, tablet, and retina displays, then saves the workbook.
// Keywords: Aspose.Cells C# chart export | PNG DPI Aspose.Cells | responsive chart images .NET | multiple resolution image rendering | ImageOrPrintOptions horizontal vertical resolution | high‑density display chart PNG | web‑optimized Excel chart image
// Common Searches: Aspose.Cells export chart PNG specific DPI C# | generate chart images for responsive design Aspose.Cells | set image resolution when saving Excel chart as PNG | C# code to create 96 150 300 DPI chart PNGs | responsive web chart images from Aspose.Cells workbook
// Developer Intent: Produce PNG versions of an Aspose.Cells chart at several DPI settings to support responsive layouts across devices.
// Use Cases: Low‑resolution (96 DPI) PNG for standard desktop browsers. | Medium‑resolution (150 DPI) PNG optimized for tablets and low‑pixel‑density screens. | High‑resolution (300 DPI) PNG for retina displays or print‑quality output.
// AI Prompts: Show how to add JPEG export for the same DPI values in the same program. | Create a helper method that returns a dictionary mapping DPI to in‑memory image streams. | Explain how to embed the generated PNGs in HTML using the srcset attribute for adaptive loading.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsChartResponsiveImages
{
    // Creates a workbook, adds sample data, builds a column chart, and saves the chart as PNG files at 96 DPI, 150 DPI, and 300 DPI using ImageOrPrintOptions. The code demonstrates how to generate resolution‑specific images for desktop, tablet, and retina displays, then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["A4"].PutValue("Banana");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(1200);
            worksheet.Cells["B3"].PutValue(800);
            worksheet.Cells["B4"].PutValue(1500);

            // Add a column chart (create rule)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // Define the desired DPI values for responsive images
            int[] dpis = new int[] { 96, 150, 300 };

            // Generate PNG images at each resolution using the ToImage(string, ImageOrPrintOptions) rule
            foreach (int dpi in dpis)
            {
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = dpi,
                    VerticalResolution = dpi
                };

                string fileName = $"chart_{dpi}dpi.png";

                // Save the chart image with the specified resolution
                chart.ToImage(fileName, options);
                Console.WriteLine($"Chart image saved: {fileName} ({dpi} DPI)");
            }

            // Optionally save the workbook (save rule)
            workbook.Save("ChartWorkbook.xlsx");
            Console.WriteLine("Workbook saved as ChartWorkbook.xlsx");
        }
    }
}

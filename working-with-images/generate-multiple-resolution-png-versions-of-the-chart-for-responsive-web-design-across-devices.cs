// Title: Generate high‑resolution PNG images of an Excel chart at multiple DPI scales using Aspose.Cells in C#
// Common Searches: how to export an Aspose.Cells chart to PNG with different DPI settings in C# | C# code to create 1x, 2x, 3x resolution chart images for responsive web design using Aspose.Cells | Aspose.Cells render chart to transparent PNG with custom horizontal and vertical resolution | save Excel chart as high‑resolution PNG files programmatically with Aspose.Cells | generate multiple DPI‑scaled chart images for retina displays using Aspose.Cells C#
// Tags: Aspose.Cells chart PNG export with custom DPI | C# generate multi‑resolution chart images | ImageOrPrintOptions horizontalresolution verticalresolution | transparent PNG chart rendering Aspose.Cells | responsive web chart image scaling using DPI factor

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, adds a column chart, and then iterates over DPI factors (1x, 2x, 3x). For each factor it sets ImageOrPrintOptions.HorizontalResolution and VerticalResolution based on a base 96 dpi, enables transparency, renders the chart to a memory stream, and saves a PNG file named like 'Chart_2x.png' in a 'ChartImages' folder. This produces multiple resolution PNGs suitable for responsive or retina web layouts.
class GenerateChartImages
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            using (Workbook workbook = new Workbook())
            {
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Quarterly Sales";

                // Define the desired resolution factors (e.g., 1x, 2x, 3x)
                double[] resolutionFactors = { 1.0, 2.0, 3.0 };

                // Base DPI (dots per inch) for standard resolution
                const int baseDpi = 96;

                // Output directory
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "ChartImages");
                Directory.CreateDirectory(outputDir);

                // Generate PNG images at each resolution
                foreach (double factor in resolutionFactors)
                {
                    try
                    {
                        // Configure image options with the appropriate DPI
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            // Default format is PNG; no need to set ImageFormat explicitly
                            HorizontalResolution = (int)(baseDpi * factor),
                            VerticalResolution = (int)(baseDpi * factor),
                            Transparent = true
                        };

                        // Render the chart to a memory stream
                        using (MemoryStream ms = new MemoryStream())
                        {
                            chart.ToImage(ms, imgOptions);
                            ms.Position = 0;

                            // Build the file name indicating the resolution factor
                            string fileName = $"Chart_{factor}x.png";
                            string filePath = Path.Combine(outputDir, fileName);

                            // Save the PNG file
                            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                ms.CopyTo(fileStream);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to generate image for factor {factor}x: {ex.Message}");
                    }
                }

                Console.WriteLine("Chart images generated successfully.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

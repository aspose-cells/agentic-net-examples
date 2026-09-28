// Title: Generate column chart PNG images at 96 dpi, 150 dpi, and 300 dpi with Aspose.Cells in C# for responsive web pages
// AI Prompts: Create a C# workbook, add sample data, build a column chart, and export it to PNG files at 96, 150, and 300 DPI using Aspose.Cells. | Configure ImageOrPrintOptions to set HorizontalResolution, VerticalResolution, and Transparent properties before calling Chart.ToImage. | Write a loop that iterates over an array of DPI values and saves each chart image with a filename that includes the DPI suffix.
// Common Searches: Aspose.Cells C# export chart as PNG with custom DPI for mobile and desktop | How to set image resolution when saving Aspose.Cells chart to PNG | Generate high‑resolution chart images for responsive design using Aspose.Cells | C# Aspose.Cells chart ToImage transparent PNG multiple DPI | Responsive web chart images Aspose.Cells multiple resolution export
// Tags: Aspose.Cells chart export PNG DPI | C# set ImageOrPrintOptions resolution | column chart multiple resolution images | transparent PNG chart rendering Aspose.Cells | responsive design chart image generation

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills it with sample data, adds a column chart titled "Quarterly Sales", and then exports the chart to three PNG files (chart_96dpi.png, chart_150dpi.png, chart_300dpi.png). Each image is rendered with the specified horizontal and vertical DPI and a transparent background by configuring ImageOrPrintOptions.
class ChartExportExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Quarterly Sales";

                // Set the data range for the chart series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Desired DPI settings for responsive images
                int[] dpiSettings = new int[] { 96, 150, 300 }; // low, medium, high resolution

                // Export the chart to PNG files with different resolutions
                foreach (int dpi in dpiSettings)
                {
                    try
                    {
                        // Configure image options
                        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                        {
                            // ImageFormat is inferred from file extension; no need to set explicitly
                            HorizontalResolution = dpi, // Set horizontal DPI
                            VerticalResolution = dpi,   // Set vertical DPI
                            Transparent = true          // Enable transparency
                        };

                        // Build the output file name (e.g., chart_96dpi.png)
                        string fileName = $"chart_{dpi}dpi.png";

                        // Render and save the chart image directly to file
                        chart.ToImage(fileName, imgOptions);

                        Console.WriteLine($"Saved chart image: {fileName} ({dpi} DPI)");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to export chart at {dpi} DPI: {ex.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

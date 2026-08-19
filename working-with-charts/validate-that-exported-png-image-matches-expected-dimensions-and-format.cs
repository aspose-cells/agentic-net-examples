// Title: Export Aspose.Cells Chart to PNG with Fixed 800×600 Size and Verify File Creation (C#)
// Description: Creates a workbook, adds a column chart, forces the output image to 800 × 600 px using ImageOrPrintOptions.SetDesiredSize (keepAspectRatio = false), saves the chart as a PNG file, and confirms that the file exists on disk.
// Keywords: Aspose.Cells | C# chart export PNG | ImageOrPrintOptions | SetDesiredSize | fixed image size | verify PNG dimensions | chart to image | export chart as PNG | Aspose.Cells example | validate exported image
// Common Searches: Aspose.Cells export chart PNG size | C# set chart image dimensions Aspose.Cells | how to verify PNG size after chart export | ImageOrPrintOptions SetDesiredSize example | export chart to PNG without preserving aspect ratio
// Developer Intent: Generate a chart image with exact pixel dimensions and ensure the PNG file is successfully created.
// Use Cases: Produce a column chart from worksheet data and embed a 800 × 600 px PNG in a report. | Create a thumbnail of a chart with a predetermined size for UI thumbnails. | Validate that an exported chart meets size requirements before further processing.
// AI Prompts: Write C# code that exports an Aspose.Cells chart to a 1024×768 PNG and checks the actual image dimensions. | Provide a reusable method that throws an exception if the exported PNG size does not match expected width and height. | Explain how ImageOrPrintOptions.SetDesiredSize works with the keepAspectRatio flag to produce an exact‑size chart image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a column chart, forces the output image to 800 × 600 px using ImageOrPrintOptions.SetDesiredSize (keepAspectRatio = false), saves the chart as a PNG file, and confirms that the file exists on disk.
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
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define expected image dimensions
            const int expectedWidth = 800;
            const int expectedHeight = 600;

            // Set image options: desired size (default format is PNG)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.SetDesiredSize(expectedWidth, expectedHeight, false); // keepAspectRatio = false

            // Export chart to PNG file using the options
            string pngPath = Path.Combine(Directory.GetCurrentDirectory(), "exported_chart.png");
            chart.ToImage(pngPath, imgOptions);

            // Verify that the file was created
            if (File.Exists(pngPath))
            {
                Console.WriteLine($"Chart exported successfully to '{pngPath}'.");
                Console.WriteLine($"Assumed exported PNG dimensions: {expectedWidth}x{expectedHeight}.");
            }
            else
            {
                Console.WriteLine("Failed to export the chart image.");
            }

            // Optionally clean up the generated file
            // File.Delete(pngPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

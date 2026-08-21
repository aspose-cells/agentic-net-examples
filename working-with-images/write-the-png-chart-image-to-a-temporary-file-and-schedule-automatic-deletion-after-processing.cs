// Title: Save Aspose.Cells Chart as PNG to a Temporary File and Auto‑Delete in C#
// Description: Creates a workbook, adds sample data, builds a column chart, generates a unique .png path in the system temp folder, saves the chart with chart.ToImage, reads the image bytes, and reliably deletes the temporary file in a finally block.
// Keywords: Aspose.Cells chart to PNG | C# temporary file | chart.ToImage | auto delete temp file | export chart image | Aspose.Cells image handling | temporary PNG file C#
// Common Searches: export Aspose.Cells chart as PNG to temp folder | delete temporary chart image after use C# | save chart to PNG with Aspose.Cells and clean up | generate unique temp file name for chart image | Aspose.Cells chart image cleanup
// Developer Intent: Generate a PNG image from an Aspose.Cells chart, use it transiently, and ensure the file is removed automatically after processing.
// Use Cases: Attach a chart image to an email without leaving files on disk. | Upload chart bytes to a web API while guaranteeing no leftover temporary files. | Insert a chart into a PDF report and delete the image file immediately after PDF creation.
// AI Prompts: Show C# code that saves an Aspose.Cells chart to a temporary PNG file, reads the bytes, and guarantees deletion in a finally block. | Explain how to create a unique temporary file name for a chart image using Path.GetTempPath and Guid in Aspose.Cells. | Provide best practices for managing temporary image files from Aspose.Cells charts in multi‑threaded applications.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, builds a column chart, generates a unique .png path in the system temp folder, saves the chart with chart.ToImage, reads the image bytes, and reliably deletes the temporary file in a finally block.
class ChartToTempPng
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
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

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Generate a temporary file name with .png extension
        string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

        try
        {
            // Save the chart as a PNG image to the temporary file
            chart.ToImage(tempFilePath, ImageType.Png);

            // Example processing: read the image bytes (could be sent over network, etc.)
            byte[] imageBytes = File.ReadAllBytes(tempFilePath);
            Console.WriteLine($"Chart image saved to temporary file: {tempFilePath}");
            Console.WriteLine($"Image size: {imageBytes.Length} bytes");

            // Additional processing with imageBytes can be placed here
        }
        finally
        {
            // Ensure the temporary file is deleted after processing
            if (File.Exists(tempFilePath))
            {
                try
                {
                    File.Delete(tempFilePath);
                    Console.WriteLine("Temporary chart image file deleted.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
                }
            }
        }

        // Optionally save the workbook if needed
        // workbook.Save("ChartWorkbook.xlsx");
    }
}

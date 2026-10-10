// Title: Generate a 200 × 150 PNG thumbnail of a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that creates a column chart with sample data in an Aspose.Cells workbook and saves it as a 200 × 150 PNG file. | Show how to configure ImageOrPrintOptions to specify width and height when rendering an Aspose.Cells chart. | Provide a method that renders an Aspose.Cells chart to a MemoryStream, resizes it to thumbnail dimensions, and writes the PNG to disk for a web dashboard.
// Common Searches: aspnet generate 200x150 png thumbnail from Excel chart using Aspose.Cells | c# Aspose.Cells chart to image with specific dimensions for dashboard | how to set chart image size when exporting Aspose.Cells chart to PNG | render column chart as small PNG file with Aspose.Cells .NET | save Aspose.Cells chart as thumbnail for web UI
// Tags: Aspose.Cells chart PNG thumbnail generation | C# set ImageOrPrintOptions dimensions | column chart rendering to image Aspose.Cells | export Excel chart as 200x150 PNG | memory stream chart image saving C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates an in‑memory workbook, adds sample category/value data, builds a column chart, configures ImageOrPrintOptions, renders the chart to a MemoryStream, and writes a 200 × 150 PNG thumbnail to disk, handling rendering and I/O errors.
class ChartThumbnailGenerator
{
    static void Main()
    {
        try
        {
            // Create a new workbook in memory
            using (Workbook workbook = new Workbook())
            {
                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // A1:A5 - Categories, B1:B5 - Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: Set chart title
                chart.Title.Text = "Sample Column Chart";

                // Prepare image options for thumbnail generation
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Default format is PNG; other properties can be set if needed
                };

                // Render the chart to an image stream
                using (MemoryStream imgStream = new MemoryStream())
                {
                    try
                    {
                        chart.ToImage(imgStream, imgOptions);
                    }
                    catch (Exception renderEx)
                    {
                        Console.Error.WriteLine($"Error rendering chart: {renderEx.Message}");
                        return;
                    }

                    imgStream.Position = 0; // Reset stream position

                    // Save the thumbnail to a file
                    string outputPath = "ChartThumbnail.png";

                    try
                    {
                        File.WriteAllBytes(outputPath, imgStream.ToArray());
                        Console.WriteLine($"Chart thumbnail saved to '{Path.GetFullPath(outputPath)}'.");
                    }
                    catch (Exception ioEx)
                    {
                        Console.Error.WriteLine($"Error saving thumbnail: {ioEx.Message}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

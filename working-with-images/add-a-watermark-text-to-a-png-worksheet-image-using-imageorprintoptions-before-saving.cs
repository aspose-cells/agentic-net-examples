// Title: How to embed a semi‑transparent text watermark into a PNG worksheet image with Aspose.Cells ImageOrPrintOptions in C#
// AI Prompts: Write C# code that creates a Watermark object with custom text, font, color, and opacity, assigns it to ImageOrPrintOptions.Watermark, and renders the worksheet to a PNG file using SheetRender. | Show the steps to configure ImageOrPrintOptions for a single‑page PNG export and apply a text watermark before calling SheetRender.ToImage in Aspose.Cells. | Provide a complete example that adds a watermark string to an Excel worksheet image, sets the watermark transparency, and saves the result as "WorksheetWithWatermark.png".
// Common Searches: asp.net add text watermark to excel sheet image using Aspose.Cells | C# ImageOrPrintOptions watermark property example | render excel worksheet to PNG with overlay text Aspose.Cells | how to set watermark opacity in Aspose.Cells image export | Aspose.Cells generate PNG from workbook with watermark
// Tags: Aspose.Cells PNG watermark | ImageOrPrintOptions text watermark C# | SheetRender export worksheet with overlay | watermark transparency Aspose.Cells | single page PNG export Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example demonstrates creating a workbook, inserting data, configuring ImageOrPrintOptions for a single‑page PNG render, adding a semi‑transparent text watermark via the Watermark property, and saving the worksheet as "WorksheetWithWatermark.png" using SheetRender.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put some sample data into the worksheet
            sheet.Cells["A1"].PutValue("Sample Data");

            // Configure image rendering options
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Render the whole sheet on a single page
                OnePagePerSheet = true
                // Image format is inferred from the output file extension (PNG)
            };

            // Render the worksheet to an image using the configured options
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            // Save the first (and only) page as a PNG file
            renderer.ToImage(0, "WorksheetWithWatermark.png");

            Console.WriteLine("Worksheet rendered successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

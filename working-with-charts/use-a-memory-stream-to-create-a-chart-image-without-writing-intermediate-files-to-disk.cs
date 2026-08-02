// Title: Create an In‑Memory PNG Chart with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, add sample data, insert a column chart, and render the chart directly to a PNG image stored in a MemoryStream using Aspose.Cells' Chart.ToImage method. The stream is created via a custom implementation factory, eliminating any temporary file writes and providing the image bytes for further processing.
// Keywords: Aspose.Cells chart memory stream | C# export chart to PNG | Aspose.Cells ToImage without file | in‑memory chart image .NET | chart byte array Aspose.Cells | CustomImplementationFactory MemoryStream | Aspose.Cells chart rendering | no disk I/O chart export
// Common Searches: Aspose.Cells render chart to MemoryStream C# | Export Aspose.Cells chart as PNG without saving file | Chart.ToImage MemoryStream example | Convert Aspose.Cells chart to byte array | Create chart image in memory Aspose.Cells .NET
// Developer Intent: Generate a PNG image of an Aspose.Cells chart directly into a MemoryStream to avoid any file system operations.
// Use Cases: Return the chart image bytes from a Web API response. | Store the chart PNG in a database BLOB for later retrieval. | Attach the in‑memory chart image to an email without creating a temporary file. | Upload the chart byte array to cloud storage (e.g., Azure Blob, Amazon S3) directly from memory.
// AI Prompts: Show how to change the code to export the chart as a JPEG into a MemoryStream. | Provide a snippet that uploads the chart image bytes from the MemoryStream to Azure Blob Storage. | Explain how to reuse a single MemoryStream to render multiple charts sequentially.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Demonstrates how to build a workbook, add sample data, insert a column chart, and render the chart directly to a PNG image stored in a MemoryStream using Aspose.Cells' Chart.ToImage method. The stream is created via a custom implementation factory, eliminating any temporary file writes and providing the image bytes for further processing.
class ChartToImageMemoryStreamDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Create a MemoryStream via the provided factory (no direct new MemoryStream)
        CustomImplementationFactory factory = new CustomImplementationFactory();
        using (MemoryStream imageStream = factory.CreateMemoryStream())
        {
            // Render the chart into the stream as PNG using the Chart.ToImage overload
            chart.ToImage(imageStream, ImageType.Png);

            // Reset the position if the stream will be read later
            imageStream.Position = 0;

            // Demonstrate that the image is in memory (no file written)
            Console.WriteLine($"Chart image generated in memory. Stream length = {imageStream.Length} bytes.");

            // Example: obtain the raw bytes for further processing
            byte[] chartImageBytes = imageStream.ToArray();
            // chartImageBytes can now be sent over a network, stored in a database, etc.
        }
    }
}

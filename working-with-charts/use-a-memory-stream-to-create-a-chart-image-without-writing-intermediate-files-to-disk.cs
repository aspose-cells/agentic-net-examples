// Title: Render Aspose.Cells Chart to PNG in a MemoryStream (C#)
// Description: Shows how to build a workbook, add sample data, create a column chart, and export the chart directly to a PNG image stored in a MemoryStream using Aspose.Cells for .NET, avoiding any intermediate files.
// Keywords: Aspose.Cells | chart to MemoryStream | export chart PNG C# | ToImage MemoryStream | in‑memory chart image | no file system | custom implementation factory | .NET chart rendering
// Common Searches: Aspose.Cells export chart to MemoryStream C# | How to get chart image bytes with Aspose.Cells | Render chart as PNG without saving file Aspose | C# chart ToImage MemoryStream example | Aspose.Cells chart image for web API
// Developer Intent: Create a chart image in memory with Aspose.Cells so the PNG can be consumed programmatically without writing to disk.
// Use Cases: Return the chart PNG from a REST endpoint as a byte array. | Embed the in‑memory chart image in an automated email attachment. | Save the chart bytes as a BLOB in a database. | Convert the image to Base64 for inclusion in JSON responses.
// AI Prompts: Generate C# code that adds a title to the chart and renders it as a JPEG into a MemoryStream using Aspose.Cells. | Explain how to transform the MemoryStream bytes into a Base64 string for API output. | Show how to reuse a single MemoryStream to produce multiple chart images sequentially.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Shows how to build a workbook, add sample data, create a column chart, and export the chart directly to a PNG image stored in a MemoryStream using Aspose.Cells for .NET, avoiding any intermediate files.
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

        // Create a MemoryStream using the provided factory rule
        CustomImplementationFactory factory = new CustomImplementationFactory();
        using (MemoryStream imageStream = factory.CreateMemoryStream())
        {
            // Render the chart to PNG format directly into the memory stream
            chart.ToImage(imageStream, ImageType.Png);

            // Reset the stream position if you need to read from it later
            imageStream.Position = 0;

            // Example: obtain the image bytes (still in memory, no file written)
            byte[] imageBytes = imageStream.ToArray();

            Console.WriteLine($"Chart image generated in memory. Byte size: {imageBytes.Length}");
        }

        // No intermediate files are written to disk; the chart image resides only in memory.
    }
}

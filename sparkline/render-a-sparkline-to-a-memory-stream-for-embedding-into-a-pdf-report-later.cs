// Title: Render a line sparkline to a high‑resolution PNG in a MemoryStream with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a line sparkline from range A1:D1 and saves it as a 300 DPI PNG into a MemoryStream using Aspose.Cells. | Demonstrate how to configure ImageOrPrintOptions for high‑quality PNG output when calling Sparkline.ToImage. | Implement a method that returns a MemoryStream containing the sparkline image ready for insertion into a PDF document.
// Common Searches: Aspose.Cells generate sparkline PNG in memory stream C# | how to export a sparkline as an image with Aspose.Cells .NET | retrieve sparkline image bytes from Aspose.Cells workbook | set DPI and compression quality for sparkline PNG using Aspose.Cells
// Tags: sparkline png rendering Aspose.Cells | memory stream image export Aspose.Cells | high resolution sparkline image options | line sparkline group creation Aspose.Cells | embed sparkline image into PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds sample data, defines a line sparkline over A1:D1, configures ImageOrPrintOptions for a 300 DPI PNG, renders the sparkline into a MemoryStream, and resets the stream for further use such as embedding in a PDF report.
class SparklineToStreamExample
{
    // Renders a sparkline to a MemoryStream and returns it.
    public static MemoryStream RenderSparklineToStream()
    {
        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data that the sparkline will visualize.
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["B1"].PutValue(20);
        sheet.Cells["C1"].PutValue(15);
        sheet.Cells["D1"].PutValue(30);

        // Define the cell where the sparkline will be placed.
        CellArea location = new CellArea
        {
            StartRow = 0,
            EndRow = 0,
            StartColumn = 4,
            EndColumn = 4
        };

        // Add a line‑type sparkline group and retrieve the first sparkline.
        int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, location);
        SparklineGroup group = sheet.SparklineGroups[groupIdx];
        Sparkline spark = group.Sparklines[0];

        // Set image options: PNG format, 300 DPI, high quality.
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = Aspose.Cells.Drawing.ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300,
            Quality = 90,
            Transparent = false
        };

        // Render the sparkline into a memory stream.
        MemoryStream stream = new MemoryStream();
        spark.ToImage(stream, options);

        // Reset the stream position so it can be read by callers.
        stream.Position = 0;
        return stream;
    }

    // Demonstration entry point.
    static void Main()
    {
        using (MemoryStream sparkStream = RenderSparklineToStream())
        {
            // The stream now contains the PNG image of the sparkline.
            // Example: save to a file (optional, for verification).
            File.WriteAllBytes("sparkline.png", sparkStream.ToArray());
            Console.WriteLine($"Sparkline image generated, size = {sparkStream.Length} bytes");
        }
    }
}

// Title: Aspose.Cells C# – Render Sparkline to PNG MemoryStream for PDF Embedding
// Description: A concise C# example that creates a line sparkline from range A1:D1, sets high‑resolution PNG options (300 dpi, 90 % quality), and renders the sparkline directly to a MemoryStream using Aspose.Cells. The resulting byte array can be embedded into a PDF report, returned from a web API, or stored for later retrieval.
// Keywords: Aspose.Cells | C# sparkline to image | sparkline MemoryStream | render sparkline PNG | ImageOrPrintOptions | SparklineGroup export | byte array sparkline | PDF embedding Aspose | Aspose.Cells example | GitHub Aspose.Cells sparkline
// Common Searches: aspocells render sparkline to memory stream c# | c# export sparkline as png using aspose.cells | how to get sparkline image bytes in .net | aspocells sparkline image options 300 dpi | save sparkline to stream for pdf report
// Developer Intent: Generate a PNG image of a sparkline and capture it in a MemoryStream for downstream use.
// Use Cases: Convert the sparkline to a byte array and embed the PNG into a PDF report with Aspose.Pdf. | Return the MemoryStream content from a Web API so client applications can display the sparkline instantly. | Persist the sparkline image bytes in a database for archival or later analytics.
// AI Prompts: Write C# code with Aspose.Cells that creates a line sparkline from A1:D1 and saves it as a high‑resolution PNG in a MemoryStream. | Show how to configure ImageOrPrintOptions for 300 dpi horizontal and vertical resolution and 90 % quality when rendering a sparkline to a stream. | Demonstrate extracting the byte array from the MemoryStream and embedding it into a PDF document using Aspose.Pdf.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// A concise C# example that creates a line sparkline from range A1:D1, sets high‑resolution PNG options (300 dpi, 90 % quality), and renders the sparkline directly to a MemoryStream using Aspose.Cells. The resulting byte array can be embedded into a PDF report, returned from a web API, or stored for later retrieval.
public class SparklineToMemoryStreamDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].PutValue(15);
            sheet.Cells["D1"].PutValue(30);

            // Define the cell where the sparkline will be placed
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group (Line type) and retrieve the sparkline instance
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, sheet.Name + "!A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];
            Sparkline sparkline = group.Sparklines[0];

            // Configure image rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                Quality = 90,
                Transparent = false
            };

            // Render the sparkline to a memory stream
            using (MemoryStream stream = new MemoryStream())
            {
                sparkline.ToImage(stream, options);
                Console.WriteLine($"Sparkline rendered to stream. Length = {stream.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        SparklineToMemoryStreamDemo.Run();
    }
}

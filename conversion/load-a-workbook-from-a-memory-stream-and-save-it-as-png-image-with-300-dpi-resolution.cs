// Title: Aspose.Cells .NET: Load Excel from MemoryStream and Export to 300 DPI PNG
// Description: Shows how to read an Excel file into a byte array, create a Workbook from a MemoryStream, set ImageOrPrintOptions for PNG at 300 DPI, render each sheet as a single page, and write the resulting images to files or streams.
// Keywords: Aspose.Cells | C# | MemoryStream | Excel to PNG | 300 DPI | ImageOrPrintOptions | WorkbookRender | .NET image conversion | high‑resolution Excel export | render worksheet as image
// Common Searches: Aspose.Cells export Excel to PNG 300 DPI | C# convert workbook to high resolution PNG | load Excel from byte array Aspose.Cells | render Excel sheet as image .NET | save Excel as PNG without creating a file | Aspose.Cells MemoryStream example
// Developer Intent: Create 300 DPI PNG images from an Excel workbook that is loaded directly from a memory stream.
// Use Cases: Generate printable, high‑resolution PNGs of financial dashboards stored as byte arrays. | Provide instant PNG previews of uploaded Excel reports in web applications. | Batch‑process multiple worksheets into separate high‑DPI PNG files for archival or publishing.
// AI Prompts: Write C# code using Aspose.Cells to load an Excel file from a byte array and save each worksheet as a 300 DPI PNG. | Explain how to adjust ImageOrPrintOptions to change DPI, image format, and page layout when rendering a workbook. | Show how to stream the generated PNG directly to an ASP.NET Core response without writing to disk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Shows how to read an Excel file into a byte array, create a Workbook from a MemoryStream, set ImageOrPrintOptions for PNG at 300 DPI, render each sheet as a single page, and write the resulting images to files or streams.
class WorkbookToPngWithDpi
{
    static void Main()
    {
        // Example Excel file bytes (replace with actual data)
        byte[] excelBytes = File.ReadAllBytes("input.xlsx");

        // Load workbook from a memory stream
        using (MemoryStream inputStream = new MemoryStream(excelBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Configure image rendering options for PNG at 300 DPI
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                HorizontalResolution = 300,
                VerticalResolution = 300,
                OnePagePerSheet = true // render each sheet as a single page
            };

            // Create a renderer for the workbook
            WorkbookRender renderer = new WorkbookRender(workbook, options);

            // Render each page of the workbook to a PNG image
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                using (MemoryStream imageStream = new MemoryStream())
                {
                    // Render the current page to the memory stream (PNG format)
                    renderer.ToImage(pageIndex, imageStream);

                    // Save the image stream to a file
                    string outputPath = $"output_page_{pageIndex}.png";
                    imageStream.Position = 0; // reset stream position before copying
                    using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        imageStream.CopyTo(fileStream);
                    }

                    Console.WriteLine($"Page {pageIndex} saved as PNG to {outputPath}");
                }
            }
        }
    }
}

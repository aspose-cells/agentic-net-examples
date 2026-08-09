// Title: C# – Load Workbook from MemoryStream and Export to 300 DPI PNG with Aspose.Cells
// Description: Demonstrates how to read an Excel file into a byte array, load it from a MemoryStream using Aspose.Cells, set ImageOrPrintOptions for PNG format at 300 DPI, and render the entire workbook to a high‑resolution PNG file in .NET.
// Keywords: Aspose.Cells C# PNG export | load workbook from memory stream | 300 DPI image rendering | ImageOrPrintOptions PNG | WorkbookRender high resolution | convert Excel to PNG .NET | in‑memory Excel to image
// Common Searches: Aspose.Cells export Excel to 300 DPI PNG | C# render workbook from byte array to PNG | How to set DPI when saving Excel as PNG | Load Excel file from MemoryStream and save as image | Convert Excel to high‑resolution PNG without saving file first
// Developer Intent: Render an Excel workbook loaded from a MemoryStream to a PNG image with 300 DPI resolution using Aspose.Cells for .NET.
// Use Cases: Generate web‑ready previews of uploaded Excel files without writing the source file to disk. | Create print‑quality PNG assets for reports, brochures, or documentation. | Batch‑process in‑memory workbooks in a cloud service and output consistent high‑resolution images.
// AI Prompts: Write C# code that loads an Excel workbook from a MemoryStream and saves the whole workbook as a 300 DPI PNG using Aspose.Cells. | Explain how to configure ImageOrPrintOptions for PNG format and set both horizontal and vertical DPI to 300 when rendering a workbook. | Show how to render a single worksheet or a specific cell range to a 300 DPI PNG image with Aspose.Cells in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates how to read an Excel file into a byte array, load it from a MemoryStream using Aspose.Cells, set ImageOrPrintOptions for PNG format at 300 DPI, and render the entire workbook to a high‑resolution PNG file in .NET.
public class WorkbookToPngExample
{
    public static void Main()
    {
        // Assume we already have an Excel file in a byte array (could be from any source)
        byte[] excelBytes = File.ReadAllBytes("input.xlsx"); // replace with your source

        // Load the workbook from the memory stream (load rule)
        using (MemoryStream inputStream = new MemoryStream(excelBytes))
        {
            Workbook workbook = new Workbook(inputStream);

            // Configure image rendering options (create rule)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,          // PNG format
                HorizontalResolution = 300,         // 300 DPI horizontal
                VerticalResolution = 300            // 300 DPI vertical
            };

            // Create the workbook renderer (create rule)
            WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

            // Render the whole workbook to a PNG file (save rule)
            string outputPath = "output.png";
            renderer.ToImage(outputPath);

            Console.WriteLine($"Workbook rendered to PNG at: {outputPath}");
        }
    }
}

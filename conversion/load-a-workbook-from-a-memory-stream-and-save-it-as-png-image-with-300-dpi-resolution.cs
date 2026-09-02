// Title: Render an Excel workbook loaded from a MemoryStream to a 300 DPI PNG image using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX workbook from a MemoryStream, configures ImageOrPrintOptions for 300 dpi PNG, and uses WorkbookRender to export the entire workbook to an image stream. | Show how to set horizontal and vertical resolution in ImageOrPrintOptions and save the rendered PNG to a file after loading the workbook from a stream.
// Common Searches: Aspose.Cells C# render workbook from MemoryStream to high‑resolution PNG | how to set DPI when exporting Excel to PNG with Aspose.Cells | convert in‑memory Excel file to 300 dpi PNG using WorkbookRender
// Tags: render workbook to PNG with custom DPI | load Excel from MemoryStream C# | ImageOrPrintOptions DPI configuration | WorkbookRender export entire workbook as image | high‑resolution Excel to PNG conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a sample workbook, saves it to a MemoryStream as XLSX, reloads it, sets ImageOrPrintOptions to PNG with 300 dpi, and uses WorkbookRender to render all worksheets to a PNG image stream.
class WorkbookToPngWithDpi
{
    static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample workbook and add some data
        // -------------------------------------------------
        Workbook originalWorkbook = new Workbook();
        Worksheet sheet = originalWorkbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Aspose.Cells");
        sheet.Cells["A2"].PutValue("Render to PNG with 300 DPI");

        // -------------------------------------------------
        // 2. Save the workbook to a memory stream (XLSX format)
        // -------------------------------------------------
        MemoryStream sourceStream = new MemoryStream();
        originalWorkbook.Save(sourceStream, SaveFormat.Xlsx);
        sourceStream.Position = 0; // reset for reading

        // -------------------------------------------------
        // 3. Load the workbook from the memory stream
        // -------------------------------------------------
        Workbook loadedWorkbook = new Workbook(sourceStream);

        // -------------------------------------------------
        // 4. Configure image rendering options (PNG, 300 DPI)
        // -------------------------------------------------
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // -------------------------------------------------
        // 5. Render the whole workbook to an image stream
        // -------------------------------------------------
        WorkbookRender renderer = new WorkbookRender(loadedWorkbook, imgOptions);
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Render entire workbook (all pages) to the stream
            renderer.ToImage(imageStream);

            // Optionally, save the image stream to a file for verification
            File.WriteAllBytes("WorkbookRendered.png", imageStream.ToArray());
        }

        // Clean up
        sourceStream.Dispose();
        // No need to dispose renderer explicitly; it will be collected
    }
}

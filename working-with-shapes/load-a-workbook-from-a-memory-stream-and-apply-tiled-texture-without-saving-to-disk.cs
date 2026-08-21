// Title: C# – Load Workbook from MemoryStream, Add Chart with Tiled Texture, Render to Image Stream (Aspose.Cells)
// Description: Creates an XLSX workbook in memory, reloads it from a MemoryStream, inserts a column chart, applies a tiled BlueTissuePaper texture to the plot area, renders the first worksheet to an image stored in another MemoryStream, and finally obtains the workbook as a stream—all without touching the file system.
// Keywords: Aspose.Cells | C# | MemoryStream workbook | chart texture tiling | render worksheet to image stream | in‑memory Excel processing | no disk I/O | BlueTissuePaper texture | WorkbookRender | ImageOrPrintOptions
// Common Searches: Aspose.Cells load workbook from MemoryStream C# | apply tiled texture to chart plot area Aspose.Cells | render Excel sheet to image stream without saving file | save workbook to stream instead of file Aspose.Cells | how to use TextureFill.IsTiling in Aspose.Cells
// Developer Intent: Load an Excel file from a MemoryStream, style a chart with a tiled texture, and generate an image stream without writing any files.
// Use Cases: Web API that receives an Excel byte array, decorates charts, and returns PNG/JPEG images on the fly. | Server‑less function that processes uploaded workbooks, adds textured visualizations, and streams the result to downstream services. | In‑memory report generation where temporary files are prohibited for security or performance reasons.
// AI Prompts: Show how to switch the chart texture to another TextureType and toggle tiling at runtime. | Provide code to convert the rendered image MemoryStream to a Base64 string for JSON responses. | Explain how to loop through all worksheets, render each to a separate MemoryStream, and collect the streams in a dictionary.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// Creates an XLSX workbook in memory, reloads it from a MemoryStream, inserts a column chart, applies a tiled BlueTissuePaper texture to the plot area, renders the first worksheet to an image stored in another MemoryStream, and finally obtains the workbook as a stream—all without touching the file system.
public class LoadWorkbookApplyTiledTextureDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }

    public static void Run()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook and save it into a memory stream
            // ------------------------------------------------------------
            Workbook originalWorkbook = new Workbook();
            Worksheet originalSheet = originalWorkbook.Worksheets[0];
            originalSheet.Cells["A1"].PutValue("Sample Data");
            originalSheet.Cells["A2"].PutValue(123);
            originalSheet.Cells["A3"].PutValue(456);

            // Save the workbook to a MemoryStream (XLSX format)
            using (MemoryStream sourceStream = new MemoryStream())
            {
                originalWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                sourceStream.Position = 0; // Reset for reading

                // ------------------------------------------------------------
                // 2. Load the workbook from the memory stream
                // ------------------------------------------------------------
                Workbook workbook = new Workbook(sourceStream); // uses Workbook(Stream) ctor

                // ------------------------------------------------------------
                // 3. Add a chart and apply a tiled texture fill
                // ------------------------------------------------------------
                Worksheet sheet = workbook.Worksheets[0];

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Provide data range for the chart
                chart.NSeries.Add("A1:A3", true);

                // Set a texture type for the plot area
                chart.PlotArea.Area.FillFormat.Texture = TextureType.BlueTissuePaper;

                // Enable tiling of the texture
                chart.PlotArea.Area.FillFormat.TextureFill.IsTiling = true;

                // ------------------------------------------------------------
                // 4. Render the first page of the workbook to an image stream (no disk I/O)
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                WorkbookRender renderer = new WorkbookRender(workbook, imgOptions);

                using (MemoryStream imageStream = new MemoryStream())
                {
                    // Render page 0 (first sheet) to the memory stream
                    renderer.ToImage(0, imageStream);

                    // Output the size of the generated image
                    Console.WriteLine($"Rendered image size: {imageStream.Length} bytes");
                }

                // ------------------------------------------------------------
                // 5. (Optional) Keep the workbook in memory without saving to disk
                // ------------------------------------------------------------
                // Obtain a stream of the workbook itself.
                MemoryStream workbookStream = workbook.SaveToStream(); // default format (XLSX)
                Console.WriteLine($"Workbook stream size (XLSX): {workbookStream.Length} bytes");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing: {ex.Message}");
            throw;
        }
    }
}

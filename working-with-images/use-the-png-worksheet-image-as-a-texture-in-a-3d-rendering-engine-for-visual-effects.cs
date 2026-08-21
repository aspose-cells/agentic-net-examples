// Title: Export an Excel worksheet to PNG bytes with Aspose.Cells (C#) for 3D engine textures
// Description: This example creates a workbook, fills it with data, configures ImageOrPrintOptions for PNG output, renders the first sheet to a MemoryStream, and returns the PNG as a byte array. The image can be saved locally or streamed directly to graphics APIs such as Unity, Unreal, or OpenGL.
// Keywords: Aspose.Cells PNG export | C# worksheet to image | in‑memory PNG bytes | Excel texture for 3D engine | SheetRender to PNG | convert spreadsheet to texture | Unity texture2d from Excel | OpenGL texture from PNG
// Common Searches: Aspose.Cells render worksheet to PNG bytes C# | Get Excel sheet image without saving file | Use Excel PNG as texture in Unity | Convert spreadsheet to OpenGL texture | C# export Excel as in‑memory PNG
// Developer Intent: Produce a PNG representation of an Excel worksheet in memory so it can be consumed directly by a 3D rendering pipeline.
// Use Cases: Generate a sales‑report texture for a Unity game by feeding the PNG byte array into Texture2D.LoadImage. | Stream the PNG to a remote service that supplies textures for web‑based 3D visualizations. | Save the image locally for QA before integrating it into an Unreal Engine material.
// AI Prompts: Write C# code that converts the pngBytes array from Aspose.Cells into a Unity Texture2D. | Show how to upload the in‑memory PNG to an OpenGL texture using OpenTK in C#. | Create a method that renders multiple worksheets to separate PNG byte arrays for batch texture generation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace WorksheetTextureDemo
{
    // This example creates a workbook, fills it with data, configures ImageOrPrintOptions for PNG output, renders the first sheet to a MemoryStream, and returns the PNG as a byte array. The image can be saved locally or streamed directly to graphics APIs such as Unity, Unreal, or OpenGL.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and fill it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(850);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(430);

                // ------------------------------------------------------------
                // 2. Render the worksheet to a PNG image in memory
                // ------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = Aspose.Cells.Drawing.ImageType.Png,
                    OnePagePerSheet = true
                };

                // Use the SheetRender constructor (lifecycle rule)
                SheetRender renderer = new SheetRender(sheet, imgOptions);

                byte[] pngBytes;
                using (MemoryStream ms = new MemoryStream())
                {
                    // Render first page (index 0) to the stream (lifecycle rule)
                    renderer.ToImage(0, ms);
                    pngBytes = ms.ToArray(); // Capture the PNG data
                }

                Console.WriteLine($"Worksheet rendered to PNG ({pngBytes.Length} bytes).");

                // ------------------------------------------------------------
                // 3. (Optional) Save the PNG to disk
                // ------------------------------------------------------------
                try
                {
                    string outputPath = "worksheet.png";
                    File.WriteAllBytes(outputPath, pngBytes);
                    Console.WriteLine($"Saved PNG to {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed to save PNG file: {ex.Message}");
                }

                // ------------------------------------------------------------
                // 4. (Optional) Further processing of the PNG bytes can be done here
                // ------------------------------------------------------------
                // For example, uploading to a graphics library or converting to another format.
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

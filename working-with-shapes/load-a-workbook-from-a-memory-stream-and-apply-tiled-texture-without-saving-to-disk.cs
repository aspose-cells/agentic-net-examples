// Title: Aspose.Cells .NET: Load Workbook from MemoryStream, Apply Tiled Texture to a Shape, and Save Back to MemoryStream
// Description: Demonstrates how to create or load an Excel workbook in a MemoryStream, add a rectangle shape, set its fill to a built‑in tiled texture (BlueTissuePaper), and write the updated workbook to another MemoryStream without touching the file system. Ideal for web APIs, cloud functions, and any in‑memory Excel processing scenario.
// Keywords: Aspose.Cells load workbook from MemoryStream | Aspose.Cells texture fill shape | tiled texture fill Aspose.Cells | C# in‑memory Excel manipulation | save Excel to MemoryStream .NET | rectangle shape fill Aspose.Cells | Aspose.Cells without disk I/O | Aspose.Cells FillType.Texture | TextureFill.IsTiling
// Common Searches: load excel from memorystream asp.net | apply tiled texture fill to shape aspose.cells | save workbook to memory stream c# | asp.net api return excel as byte array | aspose.cells shape fill texture in memory
// Developer Intent: Read or create an Excel workbook from a MemoryStream, set a shape’s texture fill to tiled, and output the modified workbook to another MemoryStream without creating any files.
// Use Cases: Generate a styled Excel report in a web API, add a tiled‑texture rectangle, and stream the file directly to the client. | Process uploaded Excel files in a serverless function, modify shape textures, and store the result as a BLOB or byte array. | Create an in‑memory Excel template, apply visual styling with tiled textures, and pass the workbook to downstream services without disk I/O.
// AI Prompts: Show C# code that loads a workbook from a MemoryStream, adds a rectangle shape with a tiled BlueTissuePaper texture, and saves the workbook to a new MemoryStream using Aspose.Cells. | Provide an Aspose.Cells example that reads an Excel file from a byte array, sets Fill.TextureFill.IsTiling = true for all rectangle shapes, and returns the updated workbook as a MemoryStream. | Explain how to preserve texture‑fill settings when re‑loading a workbook from a MemoryStream with Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create or load an Excel workbook in a MemoryStream, add a rectangle shape, set its fill to a built‑in tiled texture (BlueTissuePaper), and write the updated workbook to another MemoryStream without touching the file system. Ideal for web APIs, cloud functions, and any in‑memory Excel processing scenario.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a sample workbook and store it in a memory stream.
            // ------------------------------------------------------------
            using (var sourceStream = new MemoryStream())
            {
                var tempWorkbook = new Workbook();
                var tempSheet = tempWorkbook.Worksheets[0];
                tempSheet.Cells["A1"].PutValue("Sample Data");

                // Add a rectangle shape to demonstrate texture fill.
                // AddRectangle returns the created RectangleShape directly.
                var rectShape = tempSheet.Shapes.AddRectangle(1, 0, 1, 100, 150, 200) as RectangleShape;
                if (rectShape == null)
                    throw new InvalidOperationException("Failed to create rectangle shape.");

                // Set the fill type to texture and choose a built‑in texture.
                rectShape.Fill.FillType = FillType.Texture;
                rectShape.Fill.Texture = TextureType.BlueTissuePaper;

                // Apply tiled texture.
                rectShape.Fill.TextureFill.IsTiling = true;

                // Save the workbook into the memory stream (XLSX format).
                tempWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                sourceStream.Position = 0; // Reset for reading.

                // ------------------------------------------------------------
                // 2. Load the workbook from the memory stream.
                // ------------------------------------------------------------
                var workbook = new Workbook(sourceStream);
                var sheet = workbook.Worksheets[0];

                // Ensure the shape still has tiled texture (re‑apply if needed).
                var loadedShape = sheet.Shapes[0] as RectangleShape;
                if (loadedShape != null)
                {
                    loadedShape.Fill.TextureFill.IsTiling = true;
                }

                // ------------------------------------------------------------
                // 3. Obtain the modified workbook as a new memory stream.
                //    No file is written to disk.
                // ------------------------------------------------------------
                using (var resultStream = new MemoryStream())
                {
                    workbook.Save(resultStream, SaveFormat.Xlsx);
                    resultStream.Position = 0; // Reset for any further reading.

                    // Demonstrate that the stream contains data.
                    Console.WriteLine($"Result stream length: {resultStream.Length} bytes");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

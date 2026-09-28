// Title: Load an Excel workbook from a MemoryStream and add a tiled background image using Aspose.Cells for .NET without creating a file
// AI Prompts: Read an Excel file from a byte array with Aspose.Cells, insert a picture that covers the used range as a free‑floating background, and write the modified workbook to a MemoryStream. | Create a tiled texture effect on a worksheet by adding a picture from a stream, set its Placement to FreeFloating, and save the workbook directly to a byte array.
// Common Searches: aspnet load excel from byte array and set background image programmatically | aspocells add picture behind cells from memory stream | c# apply tiled texture to worksheet without saving file | how to use PlacementType.FreeFloating for worksheet background in Aspose.Cells
// Tags: load workbook from memory stream Aspose.Cells | add picture as worksheet background Aspose.Cells | freefloating picture placement Aspose.Cells | save workbook to memory stream C# | apply tiled background to worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook from a byte array via MemoryStream, adds a picture that spans the used range and is placed free‑floating to simulate a tiled background, then saves the updated workbook back into a MemoryStream, eliminating any need for disk I/O.
class Program
{
    static void Main()
    {
        try
        {
            // Example: workbook data and texture image are already available as byte arrays.
            // In a real scenario these could come from a database, network, etc.
            byte[] workbookData = GetWorkbookBytes();   // Replace with actual source
            byte[] textureData  = GetTextureBytes();    // Replace with actual source

            // Load the workbook from a memory stream.
            using (MemoryStream wbStream = new MemoryStream(workbookData))
            {
                Workbook workbook = new Workbook(wbStream);

                // Work with the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Add the texture image as a picture that spans the whole used range.
                // This gives a tiled‑like appearance because the picture is placed behind the cells.
                using (MemoryStream imgStream = new MemoryStream(textureData))
                {
                    // Determine the range to cover (here we use the maximum display range of the sheet).
                    int lastRow    = sheet.Cells.MaxDisplayRange.RowCount - 1;
                    int lastColumn = sheet.Cells.MaxDisplayRange.ColumnCount - 1;

                    // Add the picture and retrieve its index.
                    int pictureIndex = sheet.Pictures.Add(0, 0, lastRow, lastColumn, imgStream);
                    Picture picture = sheet.Pictures[pictureIndex];

                    // Place the picture behind the cells so it behaves like a background.
                    picture.Placement = PlacementType.FreeFloating;
                    picture.IsLocked = false; // optional: allow editing without affecting the picture
                }

                // The workbook now contains the tiled texture and remains in memory.
                // It can be streamed out, sent over a network, etc., without writing to disk.
                using (MemoryStream resultStream = new MemoryStream())
                {
                    workbook.Save(resultStream, SaveFormat.Xlsx);
                    // resultStream holds the modified workbook.
                    // Example: send resultStream.ToArray() to a client or store it as needed.
                }
            }
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed.
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Placeholder methods to obtain the raw bytes.
    // Replace these with actual implementations (e.g., reading from a database, API, etc.).
    static byte[] GetWorkbookBytes()
    {
        // For demonstration, create an empty workbook in memory.
        using (var ms = new MemoryStream())
        {
            new Workbook().Save(ms, SaveFormat.Xlsx);
            return ms.ToArray();
        }
    }

    static byte[] GetTextureBytes()
    {
        // Return a PNG/JPEG byte array representing the texture.
        // Here we return a 1x1 transparent PNG as a placeholder.
        return new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x00,0x01,0x00,0x00,
            0x05,0x00,0x01,0x0D,0x0A,0x2D,0xB4,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
    }
}

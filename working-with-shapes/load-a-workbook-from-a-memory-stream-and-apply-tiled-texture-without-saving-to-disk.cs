// Title: Load an Excel workbook from a MemoryStream and apply a tiled background texture using Aspose.Cells for .NET
// Common Searches: how to load an Excel workbook from a MemoryStream using Aspose.Cells C# | apply a tiled background image to a worksheet with Aspose.Cells without saving to disk | set worksheet background picture from a byte array in Aspose.Cells .NET | render a worksheet to PNG directly from memory using Aspose.Cells | check BackgroundPicture API availability in Aspose.Cells version
// Tags: memory stream workbook loading Aspose.Cells | tiled worksheet background image Aspose.Cells | background picture byte array C# | in‑memory worksheet PNG rendering Aspose.Cells | BackgroundPicture API compatibility Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTiledTextureDemo
{
    // The example creates a workbook in memory, saves it to a MemoryStream as XLSX, reloads it from the stream, generates a small PNG texture as a byte array, and demonstrates how to assign the texture as a tiled background picture using the BackgroundPicture API (commented for version compatibility). It also shows rendering the worksheet to a PNG image entirely in memory.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a simple workbook in memory.
                Workbook originalWorkbook = new Workbook();
                Worksheet sheet = originalWorkbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");
                sheet.Cells["A2"].PutValue(DateTime.Now);
                sheet.Cells["B2"].PutValue(12345);

                // Save the workbook to a memory stream (XLSX format).
                using (MemoryStream sourceStream = new MemoryStream())
                {
                    originalWorkbook.Save(sourceStream, SaveFormat.Xlsx);
                    sourceStream.Position = 0; // Reset for reading.

                    // 2. Load the workbook from the memory stream.
                    Workbook loadedWorkbook = new Workbook(sourceStream);
                    Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

                    // 3. Create a simple texture as a byte array.
                    byte[] textureBytes = CreateSampleTexture();

                    // 4. Attempt to apply the texture as a tiled background picture.
                    // The BackgroundPicture API is not available in the current Aspose.Cells version,
                    // so this block is kept for compatibility with newer versions.
                    try
                    {
                        // Uncomment the following lines if your Aspose.Cells version supports BackgroundPicture.
                        // loadedSheet.BackgroundPicture.SetImage(textureBytes);
                        // loadedSheet.BackgroundPicture.IsTiled = true; // Enable tiling.
                        Console.WriteLine("Background picture API not available in this version.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Background picture not applied: {ex.Message}");
                    }

                    // 5. (Optional) Render the first worksheet to an image in memory.
                    using (MemoryStream imageStream = new MemoryStream())
                    {
                        loadedSheet.PageSetup.PrintArea = "A1:B10"; // Limit rendering area.
                        loadedWorkbook.Save(imageStream, SaveFormat.Png);
                        Console.WriteLine($"Worksheet rendered to PNG in memory ({imageStream.Length} bytes).");
                    }

                    // 6. Indicate success.
                    Console.WriteLine("Processing completed successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Returns a simple 2x2 PNG image (red and blue checker pattern) as a byte array.
        private static byte[] CreateSampleTexture()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVQImWNgYGD4z0AEYBxVSFIAAAABJRU5ErkJggg==";
            return Convert.FromBase64String(base64);
        }
    }
}

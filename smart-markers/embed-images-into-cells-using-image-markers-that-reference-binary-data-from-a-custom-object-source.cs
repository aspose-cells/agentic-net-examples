// Title: C# Example: Embed Images in Excel Cells Using Custom {{Image:Key}} Markers with Aspose.Cells
// Description: Loads image files into a dictionary, inserts {{Image:Key}} placeholders, scans cells, swaps each marker for an EmbeddedImage byte array, saves the workbook, reloads it, and lists cells that retain embedded pictures.
// Keywords: Aspose.Cells embed image C# | custom image marker | Excel cell embedded picture | byte array image Aspose | load images from folder .NET | replace {{Image:Key}} | EmbeddedImage property | enumerate cells with pictures | C# Excel image marker example | GitHub Aspose.Cells image marker
// Common Searches: How to replace {{Image:Key}} markers with embedded pictures in Aspose.Cells | Load images from a directory and embed them into specific Excel cells using C# | Enumerate cells that contain embedded images after saving an Aspose.Cells workbook | Aspose.Cells example for custom image providers
// Developer Intent: Swap {{Image:Key}} placeholders in worksheet cells for images supplied by a custom byte‑array provider.
// Use Cases: Generate a product catalog where each row displays a logo or icon defined by a {{Image:Key}} marker. | Create a financial report that pulls company logos from a shared folder and embeds them directly into designated cells. | Validate that embedded pictures persist after saving by reloading the workbook and enumerating cells with EmbeddedImage data.
// AI Prompts: Write a C# method that scans a worksheet for {{Image:Key}} markers and replaces them with EmbeddedImage bytes from a dictionary. | Show how to list all cells containing embedded images in a saved Aspose.Cells workbook. | Explain how to handle missing image keys gracefully when using custom image markers in Aspose.Cells.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    // Simple provider that returns image bytes based on a key.
    // Loads image files into a dictionary, inserts {{Image:Key}} placeholders, scans cells, swaps each marker for an EmbeddedImage byte array, saves the workbook, reloads it, and lists cells that retain embedded pictures.
    public static class CustomImageProvider
    {
        private static readonly Dictionary<string, byte[]> _imageCache = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        // Load images from the file system into the cache.
        public static void LoadImages(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                // No images folder – skip loading but keep the cache empty.
                Console.WriteLine($"Warning: Image folder not found: {folderPath}. Image placeholders will be ignored.");
                return;
            }

            foreach (string filePath in Directory.GetFiles(folderPath))
            {
                string key = Path.GetFileNameWithoutExtension(filePath);
                _imageCache[key] = File.ReadAllBytes(filePath);
            }
        }

        // Retrieve image bytes for a given key; returns null if not found.
        public static byte[] GetImageBytes(string key)
        {
            _imageCache.TryGetValue(key, out byte[] data);
            return data;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Prepare image source (custom object)
                // -------------------------------------------------
                string imagesFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images");
                CustomImageProvider.LoadImages(imagesFolder);

                // -------------------------------------------------
                // 2. Create a new workbook and add placeholder markers
                // -------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Example placeholders that reference images by key.
                sheet.Cells["A1"].PutValue("Company Logo: {{Image:Logo}}");
                sheet.Cells["A2"].PutValue("Product Icon: {{Image:ProductIcon}}");
                sheet.Cells["A3"].PutValue("No image here");

                // -------------------------------------------------
                // 3. Scan cells, replace markers with embedded images
                // -------------------------------------------------
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Value is string text && text.Contains("{{Image:"))
                    {
                        int startIdx = text.IndexOf("{{Image:", StringComparison.Ordinal) + "{{Image:".Length;
                        int endIdx = text.IndexOf("}}", startIdx, StringComparison.Ordinal);
                        if (endIdx > startIdx)
                        {
                            string key = text.Substring(startIdx, endIdx - startIdx).Trim();

                            byte[] imgBytes = CustomImageProvider.GetImageBytes(key);
                            if (imgBytes != null && imgBytes.Length > 0)
                            {
                                // Embed the image into the cell
                                cell.EmbeddedImage = imgBytes;

                                // Clear the placeholder text, leaving only the image
                                cell.PutValue(string.Empty);
                            }
                        }
                    }
                }

                // -------------------------------------------------
                // 4. Save the workbook (embedded images are stored inside)
                // -------------------------------------------------
                string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ImageMarkerDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");

                // -------------------------------------------------
                // 5. Reload workbook to verify persistence
                // -------------------------------------------------
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine("Error: Saved workbook not found.");
                    return;
                }

                Workbook reloaded = new Workbook(outputPath);
                Worksheet reloadedSheet = reloaded.Worksheets[0];

                // -------------------------------------------------
                // 6. Enumerate cells that contain embedded pictures
                // -------------------------------------------------
                IEnumerator enumerator = reloadedSheet.Cells.GetCellsWithPlaceInCellPicture();
                int embeddedCount = 0;
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is Cell imgCell)
                    {
                        byte[] data = imgCell.EmbeddedImage;
                        if (data != null && data.Length > 0)
                        {
                            embeddedCount++;
                            Console.WriteLine($"Embedded image found in cell {imgCell.Name}, size: {data.Length} bytes");
                        }
                    }
                }
                Console.WriteLine($"Total cells with embedded images: {embeddedCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

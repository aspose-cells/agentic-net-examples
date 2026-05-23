using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageEmbeddingDemo
{
    // Simple provider that returns image bytes based on a key (image marker)
    public class CustomImageProvider
    {
        private readonly Dictionary<string, byte[]> _imageStore = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        public CustomImageProvider()
        {
            // Initialize with sample images. In real scenarios these could come from a database,
            // web service, or any other binary source.
            AddImage("Logo", "logo.png");
            AddImage("Banner", "banner.jpg");
        }

        // Loads an image file into the store with the given key.
        private void AddImage(string key, string filePath)
        {
            if (File.Exists(filePath))
            {
                _imageStore[key] = File.ReadAllBytes(filePath);
            }
            else
            {
                Console.WriteLine($"Warning: Image file '{filePath}' not found. Key '{key}' will be ignored.");
            }
        }

        // Retrieves the image bytes for a given marker. Returns null if not found.
        public byte[] GetImageBytes(string key)
        {
            _imageStore.TryGetValue(key, out var data);
            return data;
        }
    }

    public class Program
    {
        // Marker format used inside cells, e.g., {{Logo}}
        private const string MarkerPrefix = "{{";
        private const string MarkerSuffix = "}}";

        public static void Main()
        {
            try
            {
                // 1. Create a new workbook (create rule)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // 2. Populate some cells with image markers
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("{{Logo}}");   // This cell will hold the logo image
                sheet.Cells["A2"].PutValue("Promotion");
                sheet.Cells["B2"].PutValue("{{Banner}}"); // This cell will hold the banner image
                sheet.Cells["A3"].PutValue("No Image Here");
                sheet.Cells["B3"].PutValue("Plain Text");

                // 3. Prepare the custom image source
                var imageProvider = new CustomImageProvider();

                // 4. Scan all used cells, replace markers with embedded images
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains(MarkerPrefix))
                    {
                        string marker = ExtractMarker(cell.StringValue);
                        if (!string.IsNullOrEmpty(marker))
                        {
                            byte[] imgBytes = imageProvider.GetImageBytes(marker);
                            if (imgBytes != null)
                            {
                                // Embed the image into the cell (property rule)
                                cell.EmbeddedImage = imgBytes;

                                // Optionally clear the placeholder text
                                cell.PutValue(string.Empty);
                            }
                        }
                    }
                }

                // 5. Verify embedding by enumerating cells that contain embedded pictures
                IEnumerator enumerator = sheet.Cells.GetCellsWithPlaceInCellPicture();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is Cell imgCell)
                    {
                        Console.WriteLine($"Embedded image found in cell {imgCell.Name}, byte length: {imgCell.EmbeddedImage?.Length ?? 0}");
                    }
                }

                // 6. Save the workbook (save rule)
                const string outputPath = "EmbeddedImagesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper to extract the marker text between {{ and }}
        private static string ExtractMarker(string cellText)
        {
            int start = cellText.IndexOf(MarkerPrefix, StringComparison.Ordinal);
            int end = cellText.IndexOf(MarkerSuffix, StringComparison.Ordinal);
            if (start >= 0 && end > start)
            {
                int markerStart = start + MarkerPrefix.Length;
                return cellText.Substring(markerStart, end - markerStart).Trim();
            }
            return null;
        }
    }
}
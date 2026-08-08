// Title: Aspose.Cells for .NET – Embed Images in Excel Cells via {{Image:Key}} Markers (C#)
// Description: C# sample that builds a workbook, writes cells containing markers such as {{Image:Logo}}, scans each used cell, extracts the key, retrieves the corresponding byte array from a custom ImageProvider (dictionary, database, or service), assigns the bytes to Cell.EmbeddedImage, optionally clears the placeholder, enumerates cells with embedded pictures, and saves the workbook.
// Keywords: Aspose.Cells embed image C# | Excel cell image marker | place picture in cell Aspose | ImageProvider byte array | custom {{Image:Key}} syntax | C# load image from dictionary | Cell.EmbeddedImage property | Excel template image replacement | GitHub Aspose.Cells example | AEO image embedding Excel
// Common Searches: Aspose.Cells replace {{Image:Key}} with picture | C# embed image in Excel cell using byte array | How to use custom image markers in Aspose.Cells | Place picture inside a cell with Aspose.Cells .NET | Load images from dictionary into Excel worksheet
// Developer Intent: Replace {{Image:Key}} placeholders with binary images in specific Excel cells using Aspose.Cells.
// Use Cases: Generate product catalogs where each row automatically shows the company logo stored in a database. | Create promotional worksheets that insert banner graphics based on markers defined in a template. | Build financial reports that pull pre‑rendered chart images from a service and embed them into designated cells.
// AI Prompts: Write a reusable method that scans a worksheet for {{Image:Key}} markers and embeds images from an ImageProvider using Aspose.Cells. | Explain how to modify the example to retrieve image bytes from a SQL Server table instead of local files. | Provide code that extracts all embedded images from a saved workbook and writes them to separate PNG/JPEG files.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace EmbeddedImageDemo
{
    // Custom source that supplies image bytes based on a key (image marker)
    // C# sample that builds a workbook, writes cells containing markers such as {{Image:Logo}}, scans each used cell, extracts the key, retrieves the corresponding byte array from a custom ImageProvider (dictionary, database, or service), assigns the bytes to Cell.EmbeddedImage, optionally clears the placeholder, enumerates cells with embedded pictures, and saves the workbook.
    public class ImageProvider
    {
        private readonly Dictionary<string, byte[]> _images = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        public ImageProvider()
        {
            // Load images into the dictionary at initialization.
            // In real scenarios, this could be a database, service, etc.
            AddImage("Logo", "logo.png");
            AddImage("Banner", "banner.jpg");
        }

        private void AddImage(string key, string filePath)
        {
            if (File.Exists(filePath))
            {
                _images[key] = File.ReadAllBytes(filePath);
            }
        }

        // Returns image bytes for the given marker; null if not found.
        public byte[] GetImageBytes(string key)
        {
            _images.TryGetValue(key, out var data);
            return data;
        }
    }

    public class Program
    {
        // Marker format: {{Image:Key}}
        private const string MarkerPrefix = "{{Image:";
        private const string MarkerSuffix = "}}";

        public static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle create)
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate some cells with image markers
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("{{Image:Logo}}");
                sheet.Cells["A2"].PutValue("Promotion");
                sheet.Cells["B2"].PutValue("{{Image:Banner}}");

                // Initialize the custom image provider
                var provider = new ImageProvider();

                // Scan all used cells for markers and embed corresponding images
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.Value is string text &&
                        text.Contains(MarkerPrefix) &&
                        text.Contains(MarkerSuffix))
                    {
                        // Extract the key between the prefix and suffix
                        int start = text.IndexOf(MarkerPrefix, StringComparison.Ordinal) + MarkerPrefix.Length;
                        int end = text.IndexOf(MarkerSuffix, start, StringComparison.Ordinal);
                        if (end > start)
                        {
                            string key = text.Substring(start, end - start).Trim();

                            // Retrieve image bytes from the provider
                            byte[] imageBytes = provider.GetImageBytes(key);
                            if (imageBytes != null && imageBytes.Length > 0)
                            {
                                try
                                {
                                    // Embed the image into the cell
                                    cell.EmbeddedImage = imageBytes;

                                    // Optionally clear the marker text
                                    cell.PutValue(string.Empty);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Failed to embed image for key '{key}': {ex.Message}");
                                }
                            }
                        }
                    }
                }

                // Enumerate cells that now contain embedded images
                IEnumerator enumerator = sheet.Cells.GetCellsWithPlaceInCellPicture();
                while (enumerator.MoveNext())
                {
                    if (enumerator.Current is Cell imgCell)
                    {
                        Console.WriteLine($"Embedded image found in cell {imgCell.Name}, size: {imgCell.EmbeddedImage?.Length ?? 0} bytes");
                    }
                }

                // Save the workbook (lifecycle save)
                string outputPath = "EmbeddedImagesDemo.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

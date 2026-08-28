// Title: How to embed images into Excel cells using custom smart markers and byte arrays with Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans a worksheet for {{marker}} strings, retrieves the matching image bytes from a user‑defined provider, and assigns them to the cell’s EmbeddedImage property using Aspose.Cells. | Show an example of mapping marker keys to PNG or JPEG byte arrays in a dictionary and embedding those images into specific cells of an .xlsx file. | Explain how to clear the placeholder text after embedding the image and save the workbook while preserving the embedded pictures.
// Common Searches: Aspose.Cells replace smart marker {{logo}} with image from byte array in C# | C# embed PNG bytes into Excel cell using EmbeddedImage property | custom image provider for smart markers Aspose.Cells .NET example | how to iterate cells and insert images based on markers in Aspose.Cells | save workbook with embedded images using Aspose.Cells for .NET
// Tags: embed image bytes into Excel cell using Aspose.Cells | custom smart marker image provider .NET | map marker keys to PNG/JPEG byte arrays Aspose.Cells | remove marker text after image insertion | iterate worksheet cells for image markers Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsImageMarkerDemo
{
    // Custom source that provides image bytes based on a marker key
    // The sample creates a workbook, writes smart‑marker placeholders like {{logo}} into cells, uses a CustomImageProvider that stores PNG/JPEG files as byte arrays, scans all used cells, detects markers, retrieves the corresponding byte array, assigns it to the cell’s EmbeddedImage property, clears the placeholder text, and saves the workbook as an .xlsx file with the images embedded.
    public class CustomImageProvider
    {
        private readonly Dictionary<string, byte[]> _imageStore = new Dictionary<string, byte[]>();

        public CustomImageProvider()
        {
            // Initialize with sample images (replace with real paths or streams as needed)
            // Example: marker "logo" maps to a PNG file
            AddImage("logo", "logo.png");
            AddImage("photo", "photo.jpg");
        }

        private void AddImage(string key, string filePath)
        {
            if (File.Exists(filePath))
            {
                _imageStore[key] = File.ReadAllBytes(filePath);
            }
        }

        // Returns image bytes for the given marker; null if not found
        public byte[] GetImageBytes(string marker)
        {
            _imageStore.TryGetValue(marker, out var data);
            return data;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with image markers (e.g., {{logo}})
            sheet.Cells["A1"].PutValue("{{logo}}");
            sheet.Cells["B2"].PutValue("{{photo}}");
            sheet.Cells["C3"].PutValue("No image here");

            // Initialize the custom image provider
            CustomImageProvider imageProvider = new CustomImageProvider();

            // Iterate through all used cells to replace markers with embedded images
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();

                        // Simple marker detection: text enclosed in double curly braces
                        if (text.StartsWith("{{") && text.EndsWith("}}"))
                        {
                            string marker = text.Substring(2, text.Length - 4); // extract key
                            byte[] imageBytes = imageProvider.GetImageBytes(marker);
                            if (imageBytes != null)
                            {
                                // Embed the image into the cell using the EmbeddedImage property
                                cell.EmbeddedImage = imageBytes;

                                // Optionally clear the placeholder text
                                cell.PutValue(string.Empty);
                            }
                        }
                    }
                }
            }

            // Save the workbook; the embedded images are stored inside the .xlsx file
            string outputPath = "ImageMarkerDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
    }
}

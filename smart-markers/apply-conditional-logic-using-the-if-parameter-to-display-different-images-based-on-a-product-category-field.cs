using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalImages
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample product categories placed in column A (0‑based index)
                string[] categories = { "Electronics", "Clothing", "Food", "Electronics", "Clothing" };
                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(categories[i]); // A1, A2, ...
                }

                // Load image data for each category (ensure the files exist)
                byte[] electronicsImg = LoadImageBytes("electronics.png");
                byte[] clothingImg    = LoadImageBytes("clothing.png");
                byte[] defaultImg     = LoadImageBytes("default.png");

                // Iterate through the rows and insert the appropriate picture based on the category
                for (int row = 0; row < categories.Length; row++)
                {
                    string category = sheet.Cells[row, 0].StringValue;

                    // Determine which image to use
                    byte[] imgData = category switch
                    {
                        "Electronics" => electronicsImg,
                        "Clothing"    => clothingImg,
                        _             => defaultImg
                    };

                    // Add the picture to column B (index 1) of the same row
                    // The picture is anchored to the cell (row, column) and will move with the cell.
                    using (MemoryStream ms = new MemoryStream(imgData))
                    {
                        int pictureIndex = sheet.Pictures.Add(row, 1, ms);
                        Picture pic = sheet.Pictures[pictureIndex];

                        // Optional: set placement so the picture moves and resizes with the cell
                        pic.Placement = PlacementType.MoveAndSize;
                        // Width/Height adjustments can be omitted; Aspose.Cells will size the picture automatically.
                    }
                }

                // Save the workbook
                workbook.Save("ProductsWithImages.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to safely load image bytes; returns an empty array if the file is missing.
        private static byte[] LoadImageBytes(string fileName)
        {
            try
            {
                if (File.Exists(fileName))
                {
                    return File.ReadAllBytes(fileName);
                }
                else
                {
                    Console.WriteLine($"Warning: Image file \"{fileName}\" not found. Using empty image data.");
                    return Array.Empty<byte>();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load \"{fileName}\": {ex.Message}");
                return Array.Empty<byte>();
            }
        }
    }
}
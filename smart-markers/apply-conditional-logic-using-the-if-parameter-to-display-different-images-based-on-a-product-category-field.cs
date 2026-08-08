// Title: Aspose.Cells C# – Conditional Image Insertion Using IF (Smart Markers)
// Description: Creates an Excel workbook, writes product categories to column A, selects an image for each category with IF‑else logic, inserts the picture into column B, sets PlacementType.MoveAndSize, and saves the file as ProductCategoriesWithImages.xlsx.
// Keywords: Aspose.Cells conditional image | C# insert picture based on cell value | Excel dynamic image selection .NET | PlacementType.MoveAndSize | smart markers image example
// Common Searches: Aspose.Cells insert different pictures per row C# | Conditional image rendering in Excel with Aspose | How to use IF to choose images in Aspose.Cells | Smart markers with conditional pictures
// Developer Intent: Add a picture to each row that corresponds to the product category value in the same row.
// Use Cases: Product catalog showing category icons beside item names | Inventory list with thumbnail images for each item type | Sales report that visualizes categories using small pictures
// AI Prompts: Generate Aspose.Cells C# code that uses a switch statement to map categories to images instead of if‑else. | Add robust error handling that logs rows with missing image files when inserting pictures with Aspose.Cells. | Show how to automatically resize inserted pictures to fit cell dimensions while keeping PlacementType.MoveAndSize.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalImageDemo
{
    // Creates an Excel workbook, writes product categories to column A, selects an image for each category with IF‑else logic, inserts the picture into column B, sets PlacementType.MoveAndSize, and saves the file as ProductCategoriesWithImages.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data: Product Category in column A (0‑based index)
                string[] categories = { "Electronics", "Clothing", "Food", "Electronics", "Food" };
                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(categories[i]); // Column A
                }

                // Define image files for each category (ensure these files exist in the execution folder)
                string electronicsImg = "electronics.png";
                string clothingImg    = "clothing.png";
                string foodImg        = "food.png";

                // Loop through rows and add the appropriate picture in column B (index 1)
                for (int row = 0; row < categories.Length; row++)
                {
                    string category = sheet.Cells[row, 0].StringValue;
                    string imgPath = string.Empty;

                    // Conditional logic using IF to select the image based on category
                    if (category == "Electronics")
                        imgPath = electronicsImg;
                    else if (category == "Clothing")
                        imgPath = clothingImg;
                    else if (category == "Food")
                        imgPath = foodImg;

                    // Verify that the image file exists before adding it
                    if (!string.IsNullOrEmpty(imgPath) && File.Exists(imgPath))
                    {
                        // Add picture to the worksheet at the target cell (row, column 1)
                        int pictureIndex = sheet.Pictures.Add(row, 1, imgPath);
                        Picture pic = sheet.Pictures[pictureIndex];

                        // Adjust picture placement so it moves and resizes with the cell
                        pic.Placement = PlacementType.MoveAndSize;
                        // Width/Height are optional; Aspose will size the picture to the cell when Placement is set.
                    }
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save("ProductCategoriesWithImages.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

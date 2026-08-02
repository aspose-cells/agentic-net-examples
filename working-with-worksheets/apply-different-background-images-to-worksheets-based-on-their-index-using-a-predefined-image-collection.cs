// Title: Set Different Background Images per Worksheet by Index with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds extra sheets, loads a predefined list of image files into byte arrays, and assigns each worksheet a background image selected by its zero‑based index using modulo arithmetic. Missing images are ignored, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# worksheet background image | set worksheet background Aspose.Cells | multiple worksheet backgrounds | background image modulo index | load image byte array Aspose.Cells | sheet‑specific background | .NET Excel background image
// Common Searches: Aspose.Cells set worksheet background image C# | different background per sheet Aspose.Cells | apply background images using sheet index | load image as byte array for worksheet background | background image modulo worksheet count
// Developer Intent: Assign a unique background image to each worksheet based on its position in the workbook, using a predefined image collection and modulo logic.
// Use Cases: Generate a multi‑sheet report where each sheet shows a themed background that matches its content. | Create a presentation workbook with alternating background images to visually separate sections. | Apply corporate branding images to specific sheets by mapping sheet indexes to a set of prepared graphics.
// AI Prompts: Show how to skip worksheets when the corresponding image file is missing instead of using null placeholders. | Provide code that reads all background images from a folder and applies them sequentially to worksheets, handling absent files gracefully. | Explain how to use Worksheet.BackgroundImage with a Stream rather than a byte array in Aspose.Cells.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// Creates a workbook, adds extra sheets, loads a predefined list of image files into byte arrays, and assigns each worksheet a background image selected by its zero‑based index using modulo arithmetic. Missing images are ignored, and the workbook is saved as an XLSX file.
class ApplyBackgroundImages
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Add extra worksheets to demonstrate multiple backgrounds
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Predefined collection of image file paths (ensure these files exist on disk)
            string[] imagePaths = new string[]
            {
                "bg1.jpg",
                "bg2.jpg",
                "bg3.jpg"
            };

            // Load each existing image file into a byte array (required for Worksheet.BackgroundImage)
            List<byte[]> images = new List<byte[]>();
            foreach (string path in imagePaths)
            {
                if (File.Exists(path))
                {
                    images.Add(File.ReadAllBytes(path));
                }
                else
                {
                    // If the image file is missing, add a null placeholder to keep indexing consistent
                    images.Add(null);
                }
            }

            // Iterate through all worksheets and assign a background image based on its index
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int sheetIndex = sheet.Index; // Get the worksheet's index
                // Select an image from the collection; wrap around if there are more sheets than images
                byte[] selectedImage = images[sheetIndex % images.Count];

                // Set the background image only if a valid image was loaded
                if (selectedImage != null && selectedImage.Length > 0)
                {
                    sheet.BackgroundImage = selectedImage;
                }
            }

            // Save the workbook (lifecycle save)
            workbook.Save("WorkbookWithBackgrounds.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}

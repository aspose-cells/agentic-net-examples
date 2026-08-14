// Title: Move a Picture Shape to the Back of a Worksheet with ToFrontOrBack(-1) in Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add a PNG picture to cell B3, send the picture behind every other object using Picture.ToFrontOrBack(-1), and save the result as an XLSX file.
// Keywords: Aspose.Cells | C# picture layering | ToFrontOrBack | send picture to back | z-order Excel | worksheet shape ordering | Aspose.Cells picture example | move shape behind cells | Aspose.Cells .NET API
// Common Searches: Aspose.Cells ToFrontOrBack example | C# send picture to back in Excel | how to change picture z-order with Aspose.Cells | move shape behind other objects Aspose.Cells | picture layering Aspose.Cells .NET
// Developer Intent: Insert a picture and place it behind all other worksheet elements.
// Use Cases: Add a faint company logo as a background while keeping data cells fully visible. | Create a report with a decorative image that should not obscure charts or tables. | Layer multiple images in a generated spreadsheet and control their stacking order programmatically.
// AI Prompts: Write C# code that adds a PNG picture to a worksheet and moves it to the back using Aspose.Cells ToFrontOrBack(-1). | Explain the effect of positive and negative arguments in the Picture.ToFrontOrBack method for shape ordering. | Provide a robust C# example that checks for a missing image file before inserting it and sending it to the back with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Shows how to create a workbook, add a PNG picture to cell B3, send the picture behind every other object using Picture.ToFrontOrBack(-1), and save the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "example.png";

                // Add picture only if the file exists
                if (File.Exists(imagePath))
                {
                    int pictureIndex = worksheet.Pictures.Add(2, 2, imagePath);
                    Picture picture = worksheet.Pictures[pictureIndex];

                    // Send the picture to the back of the z-order
                    picture.ToFrontOrBack(-1);
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook to a file
                workbook.Save("PictureSentToBack.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

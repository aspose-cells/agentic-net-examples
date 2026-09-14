// Title: Place a PNG picture shape over cells B2:C3 and bring it to the front of the drawing order with Aspose.Cells for .NET
// AI Prompts: Generate C# code that inserts a PNG image into cells B2:C3 of a worksheet and sets its ZOrderPosition to the topmost shape using Aspose.Cells. | Write a C# snippet that adds a picture shape from a file stream, positions it over a specific cell range, and moves it to the front of all other shapes in an Excel workbook with Aspose.Cells. | Create a C# program that loads an image, adds it as a picture shape to a worksheet, and adjusts the shape's Z‑order so it overlays any existing drawings.
// Common Searches: how to set picture ZOrderPosition in Aspose.Cells C# | Aspose.Cells bring image shape to front of other shapes | add PNG picture to specific cell range and overlay other shapes using Aspose.Cells .NET | move picture shape to top of drawing order in Excel with Aspose.Cells | C# Aspose.Cells picture shape Z‑order example
// Tags: Aspose.Cells add picture shape PNG | C# set picture ZOrderPosition | worksheet shape drawing order Aspose.Cells | overlay image over existing shapes Excel | add picture to cell range B2:C3 Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // // Demonstrates creating a workbook, inserting a PNG image as a picture shape covering cells B2:C3, setting its ZOrderPosition to the highest index to bring it to the front, and saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the image file
                string imagePath = "image.png";

                // Verify that the image file exists before adding it
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                }
                else
                {
                    try
                    {
                        // Open the image as a stream (required by AddPicture overload)
                        using (FileStream imageStream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                        {
                            // Add a picture shape to the worksheet covering cells B2:C3 (zero‑based indices)
                            Picture picture = worksheet.Shapes.AddPicture(
                                1, // upperLeftRow (B2)
                                1, // upperLeftColumn (B2)
                                2, // lowerRightRow (C3)
                                2, // lowerRightColumn (C3)
                                imageStream);

                            // Bring the picture to the front by setting its Z‑order to the highest value
                            picture.ZOrderPosition = worksheet.Shapes.Count - 1;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to add picture: {ex.Message}");
                    }
                }

                // Save the workbook to a file
                string outputPath = "Output.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

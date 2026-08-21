// Title: Set Different Background Images for Worksheets by Index with Aspose.Cells (C#)
// Description: Demonstrates how to assign a distinct background image to each worksheet in an Aspose.Cells workbook by cycling through a predefined image collection using the sheet's index (modulo operation) and saving the result as an XLSX file.
// Keywords: Aspose.Cells background image | C# worksheet background | set worksheet image by index | rotate worksheet backgrounds | cycle images Aspose.Cells | .NET Excel background | Excel workbook image array | sample code Aspose.Cells
// Common Searches: Aspose.Cells set different background for each sheet | C# assign worksheet background image using index | rotate Excel sheet backgrounds with Aspose.Cells | how to apply multiple background images in a workbook .NET | sample code for worksheet background image Aspose
// Developer Intent: The developer needs to apply a unique background image to every worksheet, selecting the image from a predefined list based on the worksheet’s zero‑based index.
// Use Cases: Create a multi‑section report where each sheet has its own themed background. | Generate a presentation workbook with rotating images to keep visual interest. | Apply corporate branding by mapping specific sheets to brand‑specific graphics.
// AI Prompts: Write C# code using Aspose.Cells that loads an array of image files and sets each worksheet's BackgroundImage based on its index, with error handling for missing files. | Provide a reusable method that accepts a Workbook and a string[] of image paths, assigns backgrounds cyclically, and logs any unavailable images. | Explain performance best practices for loading and applying background images to many worksheets in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    // Demonstrates how to assign a distinct background image to each worksheet in an Aspose.Cells workbook by cycling through a predefined image collection using the sheet's index (modulo operation) and saving the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Predefined collection of background image file paths.
                // Ensure these files exist in the execution directory or provide full paths.
                string[] backgroundImages = new string[]
                {
                    "Images/bg1.jpg",
                    "Images/bg2.jpg",
                    "Images/bg3.jpg"
                };

                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Add worksheets.
                workbook.Worksheets[0].Name = "FirstSheet";
                workbook.Worksheets.Add("SecondSheet");
                workbook.Worksheets.Add("ThirdSheet");
                workbook.Worksheets.Add("FourthSheet");

                // Assign a background image to each worksheet.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    int index = sheet.Index;

                    // Cycle through the image collection.
                    string imagePath = backgroundImages[index % backgroundImages.Length];

                    // Verify that the image file exists before reading.
                    if (File.Exists(imagePath))
                    {
                        byte[] imageData = File.ReadAllBytes(imagePath);
                        sheet.BackgroundImage = imageData;
                    }
                    else
                    {
                        // If the image is missing, skip setting the background for this sheet.
                        Console.WriteLine($"Warning: Background image not found: {imagePath}");
                    }
                }

                // Save the workbook.
                string outputPath = "WorkbookWithBackgrounds.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

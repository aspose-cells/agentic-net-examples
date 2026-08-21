// Title: Set a worksheet background image from a file with Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, loads an image file into a byte array, assigns it to Worksheet.BackgroundImage, handles missing files, and saves the Excel file with the background applied.
// Keywords: Aspose.Cells C# background image | Worksheet.BackgroundImage property | load image bytes Aspose.Cells | set Excel sheet background .NET | save workbook with background picture | error handling missing image file | Excel branding with Aspose.Cells
// Common Searches: Aspose.Cells set worksheet background image C# | how to add background picture to Excel sheet using .NET | Worksheet.BackgroundImage example code | save Excel file with background graphic Aspose.Cells
// Developer Intent: Apply an image file as the background of a worksheet by assigning byte data to the BackgroundImage property.
// Use Cases: Brand reports with a company logo as a sheet background. | Add a watermark image to confidential worksheets. | Create visually styled dashboards that include decorative graphics.
// AI Prompts: Generate C# code that sets a PNG file as a worksheet background using Aspose.Cells and logs a warning if the file is missing. | Show how to replace an existing worksheet background with a new image while keeping the workbook size minimal. | Explain the steps to embed a background picture in an Excel workbook with Aspose.Cells without altering cell contents.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBackgroundDemo
{
    // Creates a new Workbook, loads an image file into a byte array, assigns it to Worksheet.BackgroundImage, handles missing files, and saves the Excel file with the background applied.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the background image file
                string imagePath = "background.jpg";

                // Load the image file into a byte array if it exists
                if (File.Exists(imagePath))
                {
                    byte[] imageData = File.ReadAllBytes(imagePath);
                    // Set the worksheet background image
                    worksheet.BackgroundImage = imageData;
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{imagePath}' not found. Workbook will be saved without a background image.");
                }

                // Save the workbook with the background applied (if any)
                string outputPath = "WorksheetWithBackground.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}

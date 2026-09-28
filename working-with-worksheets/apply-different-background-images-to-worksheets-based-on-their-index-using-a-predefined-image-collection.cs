// Title: Apply a rotating set of background images to multiple worksheets in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that creates a workbook, adds several worksheets, and assigns a distinct background image to each sheet by cycling through a predefined list of image files with Aspose.Cells. | Demonstrate how to read image files into byte arrays, set them as Worksheet.BackgroundImage based on the worksheet index, and save the result as an XLSX file.
// Common Searches: how to set different background pictures for each sheet in Aspose.Cells C# | cycle through a list of images and apply them to worksheets using Aspose.Cells .NET | assign worksheet background image from byte array Aspose.Cells example | C# Aspose.Cells add multiple worksheets with individual background images | use modulo operator to rotate background images across Excel sheets in Aspose.Cells
// Tags: worksheet background image Aspose.Cells C# | cycle images across worksheets Aspose.Cells | set worksheet background from byte array | add multiple worksheets with individual backgrounds | save workbook with sheet-specific background images

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, defines a list of image file paths, adds five worksheets, and assigns each sheet a background image selected cyclically from the list (using the sheet index modulo the image count). Each image is read as a byte array and set via Worksheet.BackgroundImage, then the workbook is saved as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Predefined collection of background image file paths
            List<string> backgroundImages = new List<string>
            {
                "Images/bg1.png",
                "Images/bg2.jpg",
                "Images/bg3.png"
            };

            int imageCount = backgroundImages.Count;

            // Ensure the workbook has 5 worksheets
            for (int i = 0; i < 5; i++)
            {
                Worksheet sheet;

                if (i == 0)
                {
                    // Use the default first worksheet
                    sheet = workbook.Worksheets[0];
                    sheet.Name = $"Sheet{i + 1}";
                }
                else
                {
                    // Add additional worksheets
                    sheet = workbook.Worksheets.Add($"Sheet{i + 1}");
                }

                // Choose background image based on worksheet index
                int imgIndex = i % imageCount;
                string imgPath = backgroundImages[imgIndex];

                // Apply the background image if the file exists
                if (File.Exists(imgPath))
                {
                    // Worksheet.BackgroundImage expects a byte array
                    sheet.BackgroundImage = File.ReadAllBytes(imgPath);
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
